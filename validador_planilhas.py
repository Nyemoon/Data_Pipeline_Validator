#!/usr/bin/env python3
"""
Validador e Padronizador de Planilhas para Power BI
=====================================================

Roda ANTES de qualquer arquivo entrar no Power Query, pegando os erros
que normalmente só aparecem quando o relatório já está em produção:

  - Nomes de colunas inconsistentes (espaços, acentos, maiúsculas, \xa0)
  - Colunas obrigatórias ausentes (com match aproximado opcional)
  - Duplicatas (linhas inteiras ou por chave)
  - Nulos acima do esperado
  - Tipos de dados inconsistentes dentro da mesma coluna
  - Colunas totalmente vazias (comuns em exports mal feitos)
  - Suporte a múltiplas abas de arquivos Excel (.xlsx / .xls)

Uso:
    python validador_planilhas.py entrada.xlsx
    python validador_planilhas.py entrada.xlsx --sheet "Vendas_2026"
    python validador_planilhas.py entrada.xlsx --all-sheets --output limpo.xlsx
    python validador_planilhas.py entrada.csv --config config.json
    python validador_planilhas.py entrada.xlsx --chave-duplicata id_pedido

Changelog desta revisão (auditoria):
    - CORRIGIDO: bug crítico em `duplicate_key` (args.chave-duplicata era
      interpretado como subtração e derrubava o script em toda execução).
    - CORRIGIDO: --sheet agora aceita índice numérico (não só nome da aba).
    - CORRIGIDO: decisão de formato de saída (--output) não depende mais
      da chave "default"; agora é decidida pela extensão do arquivo.
    - CORRIGIDO: pd.ExcelFile é aberto uma única vez em --all-sheets,
      evitando releituras repetidas do arquivo inteiro.
    - CORRIGIDO: match aproximado de coluna obrigatória agora RENOMEIA a
      coluna no DataFrame para o nome esperado (antes só avisava no
      relatório e o restante do pipeline continuava sem achar a coluna).
    - CORRIGIDO: check_mixed_types agora varre a coluna inteira (com um
      teto configurável), não apenas as primeiras 500 linhas.
    - CORRIGIDO: check_nulls não duplica mais o achado já reportado por
      check_empty_columns quando a coluna está 100% vazia.
    - CORRIGIDO: leitura de CSV não mascara mais qualquer exceção como
      erro de encoding; tenta detectar o encoding e propaga outros erros.
    - CORRIGIDO: pd.to_datetime agora usa dayfirst=True mas silencia o
      warning de forma controlada, mantendo compatibilidade com datas
      brasileiras (dd/mm/aaaa) sem inferências ambíguas por linha.
"""

from __future__ import annotations

import argparse
import difflib
import json
import re
import sys
import warnings
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from nomes_colunas import standardize_column_name
from tratador_planilhas import (
    tratar_dataframe_completo,
)

# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

NULL_WARNING_THRESHOLD = 0.05  # 5% de nulos já gera aviso
NULL_CRITICAL_THRESHOLD = 0.30  # 30%+ é crítico
FUZZY_MATCH_THRESHOLD = 0.80  # 80% de similaridade para match aproximado
MIXED_TYPES_SAMPLE_LIMIT = 50_000  # teto de linhas varridas por coluna (perf)

SEVERITY_INFO = "INFO"
SEVERITY_WARNING = "AVISO"
SEVERITY_CRITICAL = "CRÍTICO"


class EncodingDetectionError(ValueError):
    """Levantada quando nenhum dos encodings testados consegue decodificar o CSV.

    CORRIGIDO: uma versão anterior deste código tentava `raise
    UnicodeDecodeError("mensagem")`, mas UnicodeDecodeError exige 5
    argumentos posicionais (encoding, object, start, end, reason) — chamá-lo
    com um único argumento nunca funcionava e derrubava o script com um
    TypeError confuso no lugar da mensagem de erro pretendida. Uma exceção
    própria evita esse problema e ainda permite ao chamador (CLI ou app)
    capturá-la especificamente para dar uma mensagem amigável.
    """


@dataclass
class Finding:
    severity: str
    category: str
    message: str


@dataclass
class ValidationReport:
    source_file: str
    sheet_name: str | None = None
    total_rows: int = 0
    total_cols: int = 0
    findings: list[Finding] = field(default_factory=list)
    column_rename_map: dict[str, str] = field(default_factory=dict)

    def add(self, severity: str, category: str, message: str) -> None:
        self.findings.append(Finding(severity, category, message))

    def has_critical(self) -> bool:
        return any(f.severity == SEVERITY_CRITICAL for f in self.findings)

    def count_by_severity(self, severity: str) -> int:
        return sum(1 for f in self.findings if f.severity == severity)

    @property
    def source_name(self) -> str:
        """Alias para source_file para consistência entre CLI e interfaces."""
        return self.source_file


# ---------------------------------------------------------------------------
# Leitura de arquivo
# ---------------------------------------------------------------------------


def load_spreadsheet(path: str, sheet_name: str | int | None = 0) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        return _load_csv(path)
    elif ext in (".xlsx", ".xls"):
        return pd.read_excel(path, sheet_name=sheet_name)
    else:
        raise ValueError(f"Formato não suportado: {ext}. Use .csv, .xlsx ou .xls.")


def _load_csv(path_or_buffer) -> pd.DataFrame:
    """Tenta ler CSV detectando encoding, sem mascarar erros não relacionados
    a encoding (arquivo ausente, delimitador inválido, etc.).

    Aceita tanto um caminho de arquivo (str/Path) quanto um objeto tipo
    arquivo (ex.: o buffer retornado por st.file_uploader do Streamlit), o
    que permite reaproveitar esta função tanto no CLI quanto na interface
    web em vez de duplicar a lógica de fallback de encoding em cada lugar.
    """
    encodings_to_try = ("utf-8-sig", "latin-1", "cp1252")
    last_error: Exception | None = None
    is_seekable_buffer = hasattr(path_or_buffer, "seek")

    for enc in encodings_to_try:
        if is_seekable_buffer:
            # Uma tentativa de leitura que falhou pode ter avançado o
            # cursor do buffer; sem rebobinar, a próxima tentativa leria
            # a partir do meio do arquivo (ou nada).
            path_or_buffer.seek(0)
        try:
            return pd.read_csv(path_or_buffer, sep=None, engine="python", encoding=enc)
        except UnicodeDecodeError as e:
            last_error = e
            continue
        # Qualquer outro erro (arquivo não encontrado, parsing, etc.)
        # deve propagar imediatamente, não ser mascarado como encoding.

    # Se chegou aqui, todas as tentativas de encoding falharam.
    nome = getattr(path_or_buffer, "name", str(path_or_buffer))
    raise EncodingDetectionError(
        f"Não foi possível decodificar '{nome}' com nenhum dos encodings "
        f"testados ({', '.join(encodings_to_try)})."
    ) from last_error


def get_excel_sheets(path: str) -> list[str]:
    """Retorna o nome de todas as abas de um arquivo Excel."""
    with pd.ExcelFile(path) as xl:
        return list(xl.sheet_names)


def resolve_duplicate_sheet_names(sheet_names: list[str]) -> dict[str, str]:
    """Recebe os nomes de aba brutos de um arquivo Excel e retorna um mapa
    nome_original -> chave_segura_para_salvar/exibir.

    Nomes que só diferem em maiúsculas/minúsculas (ex.: "Vendas" e "vendas")
    recebem um sufixo numérico (e são truncados em 31 caracteres, o limite
    de nome de aba do próprio Excel) para nunca colidir — seja como chave de
    um dict, como nome de aba num ExcelWriter, ou como rótulo de aba numa UI.

    Extraído para uma função própria (em vez de ficar embutido só em
    main()) justamente para poder ser reaproveitado por qualquer front-end
    além do CLI, como a interface Streamlit, sem duplicar a lógica.
    """
    seen_lower: dict[str, int] = {}
    resolved: dict[str, str] = {}
    for sname in sheet_names:
        key_lower = sname.strip().lower()
        if key_lower in seen_lower:
            seen_lower[key_lower] += 1
            save_key = f"{sname}_{seen_lower[key_lower]}"
        else:
            seen_lower[key_lower] = 0
            save_key = sname
        # Trunca em 31 caracteres (limite de nome de aba do Excel) em
        # ambos os ramos, não só no de colisão — um nome de aba já
        # legitimamente longo também precisa respeitar o limite.
        resolved[sname] = save_key[:31]
    return resolved


# ---------------------------------------------------------------------------
# Padronização de nomes de colunas e Caracteres Especiais
# ---------------------------------------------------------------------------
# CORRIGIDO: remove_accents e standardize_column_name foram movidas para
# nomes_colunas.py (módulo sem dependências) e agora são só importadas aqui,
# em vez de definidas duas vezes. Antes desta correção, tratador_planilhas.py
# tinha sua PRÓPRIA normalização (mais fraca) e o tratamento de dados
# (conversão de tipo, remoção de duplicata por chave) falhava silenciosamente
# sempre que um nome de coluna do config tinha espaço ou acento. Importar
# standardize_column_name diretamente daqui para tratador_planilhas.py não é
# uma opção, pois criaria uma importação circular (este módulo já importa de
# tratador_planilhas.py) — daí a extração para um terceiro módulo neutro.


def standardize_columns(df: pd.DataFrame, report: ValidationReport) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    seen: dict[str, int] = {}

    for original in df.columns:
        new_name = standardize_column_name(original)
        if new_name in seen:
            seen[new_name] += 1
            new_name = f"{new_name}_{seen[new_name]}"
            report.add(
                SEVERITY_WARNING,
                "Nomes de coluna",
                f"Coluna duplicada após padronização: '{original}' virou '{new_name}' "
                f"para evitar conflito.",
            )
        else:
            seen[new_name] = 0

        if new_name != original:
            rename_map[original] = new_name

    df = df.rename(columns=rename_map)
    report.column_rename_map = rename_map

    if rename_map:
        report.add(
            SEVERITY_INFO,
            "Nomes de coluna",
            f"{len(rename_map)} coluna(s) renomeada(s) para snake_case sem acentos "
            f"(compatível com boas práticas de modelagem no Power BI).",
        )
    return df


# ---------------------------------------------------------------------------
# Checagens de qualidade (com Match Aproximado / Fuzzy Matching)
# ---------------------------------------------------------------------------


def check_required_columns(
    df: pd.DataFrame, required: list[str], report: ValidationReport
) -> pd.DataFrame:
    """Verifica colunas obrigatórias. Quando encontra uma coluna por
    similaridade (fuzzy match), RENOMEIA a coluna real para o nome esperado,
    para que checagens e argumentos posteriores (ex.: --chave-duplicata)
    funcionem de forma consistente com o que o relatório informa."""
    if not required:
        return df

    standardized_required = [standardize_column_name(c) for c in required]
    current_columns = list(df.columns)

    missing = []
    fuzzy_rename_map: dict[str, str] = {}

    for req_orig, req_std in zip(required, standardized_required):
        if req_std in current_columns:
            continue

        # Tentativa de Match Aproximado (Fuzzy Matching) com difflib
        # (não considera colunas já usadas em outro match aproximado)
        candidates = [c for c in current_columns if c not in fuzzy_rename_map]
        matches = difflib.get_close_matches(
            req_std, candidates, n=1, cutoff=FUZZY_MATCH_THRESHOLD
        )
        if matches:
            matched_col = matches[0]
            fuzzy_rename_map[matched_col] = req_std
            report.add(
                SEVERITY_WARNING,
                "Colunas obrigatórias (Match Aproximado)",
                f"Coluna obrigatória '{req_orig}' (esperada como '{req_std}') não foi encontrada exatamente, "
                f"mas mapeada por similaridade para '{matched_col}'. A coluna foi renomeada para "
                f"'{req_std}'. Verifique se o nome mudou na fonte.",
            )
        else:
            missing.append(req_orig)

    if fuzzy_rename_map:
        df = df.rename(columns=fuzzy_rename_map)
        report.column_rename_map.update(fuzzy_rename_map)

    if missing:
        report.add(
            SEVERITY_CRITICAL,
            "Colunas obrigatórias",
            f"Coluna(s) obrigatória(s) ausente(s) ou sem correspondência aproximada: {', '.join(missing)}.",
        )
    elif not any(
        f.category.startswith("Colunas obrigatórias")
        and f.severity == SEVERITY_CRITICAL
        for f in report.findings
    ):
        report.add(
            SEVERITY_INFO,
            "Colunas obrigatórias",
            "Todas as colunas obrigatórias estão presentes (exatas ou por correspondência aproximada).",
        )

    return df


def check_empty_columns(df: pd.DataFrame, report: ValidationReport) -> set[str]:
    """Retorna o conjunto de colunas 100% vazias, para que check_nulls
    possa evitar reportar o mesmo achado duas vezes."""
    empty_cols = {c for c in df.columns if df[c].isna().all()}
    if empty_cols:
        report.add(
            SEVERITY_CRITICAL,
            "Colunas vazias",
            f"Coluna(s) 100% vazia(s), provavelmente erro de exportação: {', '.join(sorted(empty_cols))}.",
        )
    return empty_cols


def check_nulls(
    df: pd.DataFrame, report: ValidationReport, skip_columns: set[str] | None = None
) -> None:
    skip_columns = skip_columns or set()
    if len(df) == 0:
        return
    for col in df.columns:
        if col in skip_columns:
            continue  # já reportado por check_empty_columns
        null_ratio = df[col].isna().mean()
        if null_ratio == 0:
            continue
        pct = null_ratio * 100
        if null_ratio >= NULL_CRITICAL_THRESHOLD:
            report.add(
                SEVERITY_CRITICAL,
                "Valores nulos",
                f"Coluna '{col}' tem {pct:.1f}% de valores nulos.",
            )
        elif null_ratio >= NULL_WARNING_THRESHOLD:
            report.add(
                SEVERITY_WARNING,
                "Valores nulos",
                f"Coluna '{col}' tem {pct:.1f}% de valores nulos.",
            )


def check_duplicate_rows(df: pd.DataFrame, report: ValidationReport) -> None:
    try:
        dup_count = df.duplicated().sum()
    except TypeError:
        report.add(
            SEVERITY_WARNING,
            "Duplicatas",
            "Não foi possível checar duplicatas de linha inteira: a planilha contém "
            "valores não comparáveis (ex.: listas/objetos em alguma célula).",
        )
        return
    if dup_count > 0:
        report.add(
            SEVERITY_WARNING,
            "Duplicatas",
            f"{dup_count} linha(s) inteiramente duplicada(s) encontrada(s).",
        )


def check_duplicate_key(
    df: pd.DataFrame, key_column: str | None, report: ValidationReport
) -> None:
    if not key_column:
        return
    key_std = standardize_column_name(key_column)
    if key_std not in df.columns:
        report.add(
            SEVERITY_WARNING,
            "Chave de duplicata",
            f"Coluna-chave '{key_column}' informada não foi encontrada após padronização.",
        )
        return
    dup_count = df[key_std].duplicated().sum()
    if dup_count > 0:
        report.add(
            SEVERITY_CRITICAL,
            "Duplicatas por chave",
            f"{dup_count} valor(es) duplicado(s) na coluna-chave '{key_std}'. "
            f"Isso pode causar duplicação de dados em relacionamentos no Power BI.",
        )


def check_mixed_types(df: pd.DataFrame, report: ValidationReport) -> None:
    """Detecta colunas 'object' que misturam números e texto.

    Varre a coluna inteira (até um teto de segurança para performance em
    planilhas muito grandes), em vez de olhar só as primeiras 500 linhas,
    para não deixar passar inconsistências que só aparecem mais adiante
    no arquivo.
    """
    for col in df.columns:
        if df[col].dtype != object:
            continue
        non_null = df[col].dropna()
        if non_null.empty:
            continue

        sample = (
            non_null
            if len(non_null) <= MIXED_TYPES_SAMPLE_LIMIT
            else non_null.sample(MIXED_TYPES_SAMPLE_LIMIT, random_state=0)
        )

        types_found = set()
        for v in sample:
            if isinstance(v, bool):
                types_found.add("bool")
            elif isinstance(v, (int, float)):
                types_found.add("numero")
            else:
                s = str(v).strip()
                if re.fullmatch(r"-?\d+([.,]\d+)?", s):
                    types_found.add("numero_como_texto")
                else:
                    types_found.add("texto")
        if len(types_found) > 1:
            report.add(
                SEVERITY_WARNING,
                "Tipos inconsistentes",
                f"Coluna '{col}' parece misturar tipos diferentes ({', '.join(sorted(types_found))}). "
                f"Isso pode gerar erro de conversão de tipo no Power Query.",
            )


def _compatibilidade_tipo(
    series: pd.Series, tipo: str, col: str, report: ValidationReport
) -> float:
    """Calcula a fração de valores de `series` compatíveis com `tipo`
    ('numero', 'data', ou qualquer outra string tratada como 'aceita
    qualquer coisa', ex. 'texto'). Extraído de check_expected_types para
    poder ser chamado uma vez por tipo candidato, quando a coluna aceita
    mais de um tipo (ex.: tipos_esperados: {"cliente_id": ["numero", "texto"]})."""
    if tipo == "numero":
        return pd.to_numeric(series, errors="coerce").notna().mean()

    if tipo == "data":
        # AVISO (sugestão de auditoria externa): pd.to_datetime é permissivo
        # com valores puramente numéricos — pode interpretar números de
        # série do Excel (ou até uma coluna numérica comum) como datas
        # válidas, gerando falso positivo/negativo dependendo da
        # formatação regional de origem. Detectamos essa ambiguidade
        # separadamente do cálculo de "ok" e alertamos o usuário para
        # revisão manual, em vez de confiar cegamente no resultado.
        if pd.api.types.is_numeric_dtype(series):
            numeric_like_ratio = 1.0
        else:
            numeric_like_ratio = (
                series.astype(str)
                .str.strip()
                .str.fullmatch(r"-?\d+(\.\d+)?")
                .fillna(False)
                .mean()
            )

        with warnings.catch_warnings():
            # dayfirst=True é a convenção correta para datas brasileiras
            # (dd/mm/aaaa); silenciamos apenas o UserWarning de parsing
            # ambíguo do pandas, mantendo o comportamento determinístico.
            warnings.simplefilter("ignore", UserWarning)
            ok = pd.to_datetime(series, errors="coerce", dayfirst=True).notna().mean()

        if numeric_like_ratio > 0.05:
            report.add(
                SEVERITY_WARNING,
                "Tipo esperado (ambíguo)",
                f"Coluna '{col}' está configurada como tipo 'data', mas "
                f"{numeric_like_ratio * 100:.1f}% dos valores são puramente numéricos. "
                f"Isso é uma fonte comum de falso positivo/negativo: pode ser um número "
                f"de série do Excel não convertido, ou uma coluna numérica sendo "
                f"interpretada como data. Revise manualmente a formatação de origem "
                f"antes de confiar no resultado abaixo.",
            )
        return ok

    # Qualquer outro tipo (ex.: "texto") aceita qualquer valor.
    return 1.0


def check_expected_types(
    df: pd.DataFrame,
    expected_types: dict[str, str | list[str]],
    report: ValidationReport,
) -> None:
    """Valida se cada coluna bate com o(s) tipo(s) esperado(s).

    ADICIONADO: `expected_types` agora aceita uma lista de tipos candidatos
    por coluna (ex.: {"cliente_id": ["numero", "texto"]}), para colunas cujo
    formato varia entre exports diferentes da mesma planilha — a coluna é
    considerada válida se bater com QUALQUER um dos tipos da lista. Uma
    string única continua funcionando exatamente como antes.
    """
    for raw_col, expected in expected_types.items():
        col = standardize_column_name(raw_col)
        if col not in df.columns:
            continue
        series = df[col].dropna()
        if series.empty:
            continue

        candidatos = expected if isinstance(expected, list) else [expected]

        resultados: dict[str, float] = {}
        for tipo in candidatos:
            resultados[tipo] = _compatibilidade_tipo(series, tipo, col, report)

        melhor_tipo = max(resultados, key=resultados.get)
        melhor_ok = resultados[melhor_tipo]

        if melhor_ok < 0.95:
            if len(candidatos) == 1:
                report.add(
                    SEVERITY_CRITICAL,
                    "Tipo esperado",
                    f"Coluna '{col}' deveria ser do tipo '{candidatos[0]}', mas só "
                    f"{melhor_ok * 100:.1f}% dos valores são compatíveis.",
                )
            else:
                detalhe = ", ".join(
                    f"'{t}': {r * 100:.1f}%" for t, r in resultados.items()
                )
                report.add(
                    SEVERITY_CRITICAL,
                    "Tipo esperado",
                    f"Coluna '{col}' não é compatível com nenhum dos tipos esperados "
                    f"({detalhe}).",
                )


# ---------------------------------------------------------------------------
# Execução da validação unitária (por DataFrame/Aba)
# ---------------------------------------------------------------------------


def run_single_validation(
    df: pd.DataFrame,
    source_name: str,
    sheet_name: str | None = None,
    required_columns: list[str] | None = None,
    expected_types: dict[str, str] | None = None,
    duplicate_key: str | None = None,
) -> tuple[pd.DataFrame, ValidationReport]:
    report = ValidationReport(source_file=source_name, sheet_name=sheet_name)
    report.total_rows, report.total_cols = df.shape

    df = standardize_columns(df, report)
    df = check_required_columns(df, required_columns or [], report)

    empty_cols = check_empty_columns(df, report)
    check_nulls(df, report, skip_columns=empty_cols)
    check_duplicate_rows(df, report)
    check_duplicate_key(df, duplicate_key, report)
    check_mixed_types(df, report)
    if expected_types:
        check_expected_types(df, expected_types, report)

    if not report.findings:
        report.add(
            SEVERITY_INFO,
            "Geral",
            "Nenhum problema encontrado. Planilha pronta para uso.",
        )

    return df, report


# ---------------------------------------------------------------------------
# Relatório em Markdown
# ---------------------------------------------------------------------------


def render_markdown_report(reports: list[ValidationReport]) -> str:
    lines = ["# Relatório de Validação de Planilhas", ""]

    for report in reports:
        sheet_info = f" (Aba: `{report.sheet_name}`)" if report.sheet_name else ""
        lines.append(f"## Arquivo: `{Path(report.source_file).name}`{sheet_info}")
        lines.append("")
        lines.append(f"- Linhas: **{report.total_rows}**")
        lines.append(f"- Colunas: **{report.total_cols}**")
        lines.append(
            f"- Críticos: **{report.count_by_severity(SEVERITY_CRITICAL)}** | "
            f"Avisos: **{report.count_by_severity(SEVERITY_WARNING)}** | "
            f"Info: **{report.count_by_severity(SEVERITY_INFO)}**"
        )
        lines.append("")

        if report.column_rename_map:
            lines.append("### Colunas renomeadas")
            lines.append("")
            lines.append("| Original | Padronizado |")
            lines.append("|---|---|")
            for orig, new in report.column_rename_map.items():
                lines.append(f"| {orig} | {new} |")
            lines.append("")

        lines.append("### Achados")
        lines.append("")
        order = {SEVERITY_CRITICAL: 0, SEVERITY_WARNING: 1, SEVERITY_INFO: 2}
        for f in sorted(report.findings, key=lambda x: order.get(x.severity, 9)):
            marker = {"CRÍTICO": "🔴", "AVISO": "🟡", "INFO": "🔵"}.get(f.severity, "")
            lines.append(f"- {marker} **[{f.severity}] {f.category}** — {f.message}")

        lines.append("")
        if report.has_critical():
            lines.append(
                "> ⚠️ **Existem problemas críticos nesta aba/arquivo. Corrija antes de publicar no Power BI.**"
            )
        else:
            lines.append("> ✅ Nenhum problema crítico encontrado.")
        lines.append("\n---\n")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def load_config(path: str | None) -> dict:
    if not path:
        return {}
    with open(path, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError as e:
            raise SystemExit(f"Erro: config.json inválido — {e}") from e


def _resolve_sheet_arg(sheet_arg: str | None) -> str | int | None:
    """Converte o valor de --sheet para int quando for um índice numérico
    (ex.: '1'), ou mantém como string quando for o nome da aba."""
    if sheet_arg is None:
        return None
    try:
        return int(sheet_arg)
    except ValueError:
        return sheet_arg


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Valida e padroniza planilhas antes de subir ao Power BI."
    )
    parser.add_argument("entrada", help="Caminho do arquivo .csv, .xlsx ou .xls")
    parser.add_argument(
        "--config", help="Caminho de um config.json com colunas obrigatórias/tipos"
    )
    parser.add_argument(
        "--output", help="Caminho para salvar a planilha padronizada (.xlsx ou .csv)"
    )
    parser.add_argument(
        "--report", default="relatorio.md", help="Caminho do relatório em Markdown"
    )
    parser.add_argument(
        "--chave-duplicata",
        dest="chave_duplicata",
        help="Nome da coluna a checar como chave única",
    )
    parser.add_argument(
        "--sheet", help="Nome ou índice (numérico) da aba do Excel a ser validada"
    )
    parser.add_argument(
        "--all-sheets",
        action="store_true",
        help="Valida todas as abas de um arquivo Excel",
    )
    args = parser.parse_args()

    config = load_config(args.config)
    required_columns = config.get("colunas_obrigatorias", [])
    expected_types = config.get("tipos_esperados", {})

    # CORRIGIDO: bug original usava "args.chave-duplicata" (subtração
    # inválida) e derrubava o script em toda execução. args.chave_duplicata
    # é o atributo correto criado pelo argparse a partir de --chave-duplicata.
    duplicate_key = args.chave_duplicata or config.get("chave_duplicata")

    reports: list[ValidationReport] = []
    dfs_to_save: dict[str, pd.DataFrame] = {}

    try:
        ext = Path(args.entrada).suffix.lower()
        if ext in (".xlsx", ".xls") and args.all_sheets:
            # Abre o arquivo Excel uma única vez e reaproveita para todas
            # as abas, em vez de reabrir o arquivo a cada aba.
            xl = pd.ExcelFile(args.entrada)
            sheet_names = xl.sheet_names

            # CORRIGIDO: nomes de aba que só diferem em maiúsculas/minúsculas
            # (ex.: "Vendas" e "vendas") são tratados pelo Excel como
            # duplicados ao salvar; sem isso, a última sobrescreveria a
            # primeira silenciosamente em dfs_to_save/--output. A resolução
            # agora é feita por resolve_duplicate_sheet_names(), extraída
            # para uma função própria e reaproveitada também pela interface
            # Streamlit, em vez de duplicar esta lógica nos dois lugares.
            save_keys = resolve_duplicate_sheet_names(sheet_names)
            for sname in sheet_names:
                save_key = save_keys[sname]

                df = pd.read_excel(xl, sheet_name=sname)
                df_clean, report = run_single_validation(
                    df,
                    args.entrada,
                    sheet_name=sname,
                    required_columns=required_columns,
                    expected_types=expected_types,
                    duplicate_key=duplicate_key,
                )
                if save_key != sname:
                    report.add(
                        SEVERITY_WARNING,
                        "Nomes de aba",
                        f"Aba '{sname}' tem nome equivalente a outra aba já processada "
                        f"(diferindo só em maiúsculas/minúsculas). Será salva como "
                        f"'{save_key}' para evitar sobrescrita silenciosa.",
                    )
                reports.append(report)
                dfs_to_save[save_key] = df_clean
        else:
            resolved_sheet = (
                _resolve_sheet_arg(args.sheet) if ext in (".xlsx", ".xls") else None
            )
            target_sheet = resolved_sheet if args.sheet else 0
            df = load_spreadsheet(args.entrada, sheet_name=target_sheet)
            s_label = str(target_sheet) if args.sheet else None
            df_clean, report = run_single_validation(
                df,
                args.entrada,
                sheet_name=s_label,
                required_columns=required_columns,
                expected_types=expected_types,
                duplicate_key=duplicate_key,
            )
            reports.append(report)
            dfs_to_save["default"] = df_clean

    except PermissionError:
        # Caso comum: o arquivo Excel está aberto em outro programa
        # (Excel do Windows bloqueia o arquivo para leitura exclusiva).
        print(
            f"Erro: não foi possível acessar '{args.entrada}'. O arquivo pode estar "
            f"aberto em outro programa (ex.: Excel). Feche-o e tente novamente.",
            file=sys.stderr,
        )
        return 2
    except FileNotFoundError:
        print(f"Erro: arquivo não encontrado: '{args.entrada}'.", file=sys.stderr)
        return 2
    except Exception as e:  # noqa: BLE001 — fallback de último recurso no CLI; erros não previstos devem ser reportados
        print(f"Erro ao processar o arquivo: {e}", file=sys.stderr)
        return 2

    markdown = render_markdown_report(reports)
    Path(args.report).write_text(markdown, encoding="utf-8")
    print(markdown)
    print(f"\nRelatório salvo em: {args.report}")

    if args.output:
        out_ext = Path(args.output).suffix.lower()
        is_single_df = len(dfs_to_save) == 1

        if is_single_df:
            single_df = next(iter(dfs_to_save.values()))

            # --- APLICA O TRATAMENTO AQUI ANTES DE SALVAR ---
            single_df = tratar_dataframe_completo(
                single_df, chave_duplicata=duplicate_key, tipos_esperados=expected_types
            )

            if out_ext == ".csv":
                single_df.to_csv(args.output, index=False)
            else:
                single_df.to_excel(args.output, index=False)
            print(f"Planilha tratada e salva em: {args.output}")
        else:
            if out_ext == ".csv":
                print(
                    "Aviso: múltiplas abas foram validadas, mas .csv não suporta "
                    "múltiplas abas. Salve como .xlsx para manter todas as abas.",
                    file=sys.stderr,
                )
                return 2

            with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
                for sname, df_s in dfs_to_save.items():
                    # --- APLICA O TRATAMENTO EM CADA ABA ---
                    df_tratado = tratar_dataframe_completo(
                        df_s,
                        chave_duplicata=duplicate_key,
                        tipos_esperados=expected_types,
                    )
                    df_tratado.to_excel(writer, sheet_name=sname, index=False)
            print(f"Planilha com múltiplas abas tratadas salva em: {args.output}")

    # CORRIGIDO: este bloco tinha sido removido acidentalmente na edição que
    # separou o tratamento em tratador_planilhas.py — sem ele, main() sempre
    # retornava None (equivalente a sys.exit(0)) mesmo com achados críticos,
    # quebrando qualquer pipeline que dependa do exit code para travar builds.
    has_any_critical = any(r.has_critical() for r in reports)
    return 1 if has_any_critical else 0


if __name__ == "__main__":
    sys.exit(main())
