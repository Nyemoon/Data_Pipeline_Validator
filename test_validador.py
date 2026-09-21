"""
Testes para validador_planilhas.py
=====================================

Além dos testes unitários das funções de checagem individuais, este
arquivo inclui testes de integração de main() cobrindo especificamente
os bugs encontrados na auditoria do script — em particular o bug crítico
original em --chave-duplicata, que derrubava o script em toda execução
e que uma suíte só com testes unitários das funções internas não pegaria.
"""

from __future__ import annotations

import io
import sys
import warnings
from typing import ClassVar
from unittest.mock import patch

import openpyxl
import pandas as pd
import pytest

import validador_planilhas as vp
from nomes_colunas import standardize_column_name
from tratador_planilhas import (
    _tentar_converter,
    converter_tipos_colunas,
    remover_colunas_vazias,
    remover_duplicadas,
    tratar_dataframe_completo,
)
from validador_planilhas import (
    SEVERITY_CRITICAL,
    SEVERITY_INFO,
    SEVERITY_WARNING,
    EncodingDetectionError,
    ValidationReport,
    _compatibilidade_tipo,
    _load_csv,
    _resolve_sheet_arg,
    check_duplicate_key,
    check_duplicate_rows,
    check_empty_columns,
    check_expected_types,
    check_mixed_types,
    check_nulls,
    check_required_columns,
    get_excel_sheets,
    load_config,
    load_spreadsheet,
    render_markdown_report,
    resolve_duplicate_sheet_names,
    run_single_validation,
    standardize_columns,
)


class FakeExcelFile:
    sheet_names: ClassVar[list[str]] = ["Vendas", "vendas"]

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass


@pytest.fixture
def report() -> ValidationReport:
    """Fixture simples para evitar repetir ValidationReport(...) em cada teste."""
    return ValidationReport(source_file="teste.xlsx")


# ---------------------------------------------------------------------------
# Testes unitários — padronização de colunas
# ---------------------------------------------------------------------------

def test_standardize_columns(report):
    # Testa se remove acentos, espaços, caracteres invisíveis e converte para snake_case
    df = pd.DataFrame(columns=["  ID Pedido! ", "Data Venda\xa0", "VALOR TOTAL"])
    df_padrao = standardize_columns(df, report)

    assert list(df_padrao.columns) == ["id_pedido", "data_venda", "valor_total"]
    assert report.column_rename_map["  ID Pedido! "] == "id_pedido"


# ---------------------------------------------------------------------------
# Testes unitários — nulos e colunas vazias
# ---------------------------------------------------------------------------

def test_check_nulls(report):
    df = pd.DataFrame({
        "col_ok": [1, 2, 3, 4],
        "col_alta_nula": [None, None, None, 1]  # 75% nulo
    })
    check_nulls(df, report)

    # CORRIGIDO (revisão): a asserção original só conferia se a coluna era
    # mencionada em algum achado, sem checar a severidade. Com 75% de nulos
    # (acima do NULL_CRITICAL_THRESHOLD = 0.30), o achado tem que ser
    # especificamente CRÍTICO — não apenas "algum achado qualquer".
    criticos = [f for f in report.findings if f.severity == SEVERITY_CRITICAL]
    assert any("col_alta_nula" in f.message for f in criticos)
    assert not any("col_ok" in f.message for f in report.findings)


def test_check_nulls_skip_columns(report):
    """CORRIGIDO (auditoria): check_nulls não deve duplicar o achado de uma
    coluna já reportada como 100% vazia por check_empty_columns."""
    df = pd.DataFrame({
        "vazia": [None, None, None],
        "ok": [1, 2, 3],
    })
    empty_cols = check_empty_columns(df, report)
    check_nulls(df, report, skip_columns=empty_cols)

    categorias_vazia = [f.category for f in report.findings if "vazia" in f.message]
    # Só a checagem de "Colunas vazias" deve mencionar a coluna, não também
    # "Valores nulos" — senão o mesmo problema apareceria duas vezes.
    assert categorias_vazia.count("Valores nulos") == 0
    assert categorias_vazia.count("Colunas vazias") == 1


def test_check_empty_columns(report):
    df = pd.DataFrame({
        "preenchida": [1, 2, 3],
        "vazia": [None, None, None]
    })
    empty_cols = check_empty_columns(df, report)

    assert "vazia" in empty_cols
    assert report.has_critical()


# ---------------------------------------------------------------------------
# Testes unitários — duplicatas
# ---------------------------------------------------------------------------

def test_check_duplicate_key(report):
    df = pd.DataFrame({
        "id_pedido": [101, 102, 101],  # 101 duplicado
        "valor": [50, 60, 55]
    })
    check_duplicate_key(df, key_column="id_pedido", report=report)

    assert report.has_critical()
    assert any("id_pedido" in f.message for f in report.findings)


def test_check_duplicate_rows(report):
    df = pd.DataFrame({
        "a": [1, 2, 1],
        "b": ["x", "y", "x"],
    })  # linha 0 e 2 são inteiramente duplicadas
    check_duplicate_rows(df, report)

    assert any("Duplicatas" == f.category for f in report.findings)
    assert any("1 linha(s)" in f.message for f in report.findings)


# ---------------------------------------------------------------------------
# Testes unitários — colunas obrigatórias e match aproximado (fuzzy)
# ---------------------------------------------------------------------------

def test_check_required_columns_exact_match(report):
    df = pd.DataFrame(columns=["id_pedido", "valor"])
    df = check_required_columns(df, ["id_pedido"], report)

    assert not report.has_critical()
    assert list(df.columns) == ["id_pedido", "valor"]


def test_check_required_columns_fuzzy_rename(report):
    """CORRIGIDO (auditoria): quando uma coluna obrigatória é encontrada só
    por similaridade (ex.: erro de digitação na fonte), o match aproximado
    deve de fato RENOMEAR a coluna para o nome esperado — não apenas avisar
    no relatório e deixar o restante do pipeline sem achar a coluna."""
    df = pd.DataFrame(columns=["id_pedid", "valor"])  # falta o "o" final
    df = check_required_columns(df, ["id_pedido"], report)

    # A coluna deve ter sido renomeada de fato...
    assert "id_pedido" in df.columns
    assert "id_pedid" not in df.columns
    # ...e o relatório deve deixar claro que foi por similaridade, não exato.
    assert any(
        f.category == "Colunas obrigatórias (Match Aproximado)" for f in report.findings
    )
    # Não deve haver achado crítico de coluna ausente, já que foi resolvida.
    assert not report.has_critical()


def test_check_required_columns_missing(report):
    df = pd.DataFrame(columns=["outra_coluna"])
    check_required_columns(df, ["id_pedido"], report)

    assert report.has_critical()


# ---------------------------------------------------------------------------
# Testes unitários — tipos de dados
# ---------------------------------------------------------------------------

def test_check_mixed_types_detects_inconsistency(report):
    df = pd.DataFrame({"col": [1, "dois", 3, "quatro"]})
    check_mixed_types(df, report)

    assert any(f.category == "Tipos inconsistentes" for f in report.findings)


def test_check_mixed_types_ignores_consistent_column(report):
    df = pd.DataFrame({"col": ["a", "b", "c"]})
    check_mixed_types(df, report)

    assert not any(f.category == "Tipos inconsistentes" for f in report.findings)


def test_check_expected_types_numero_ok(report):
    df = pd.DataFrame({"valor": [1, 2, 3, 4]})
    check_expected_types(df, {"valor": "numero"}, report)

    assert not report.has_critical()


def test_check_expected_types_numero_falha(report):
    df = pd.DataFrame({"valor": ["a", "b", "c", "d"]})
    check_expected_types(df, {"valor": "numero"}, report)

    assert report.has_critical()


def test_check_expected_types_data_ambigua_gera_aviso(report):
    """ADICIONADO (sugestão de auditoria externa): uma coluna marcada como
    tipo 'data' cujos valores são puramente numéricos (ex.: número de série
    do Excel não convertido) deve gerar um AVISO de ambiguidade — não deve
    ser silenciosamente aceita nem silenciosamente rejeitada."""
    df = pd.DataFrame({"data_venda": [45832, 45833, 45834]})
    check_expected_types(df, {"data_venda": "data"}, report)

    assert any(f.category == "Tipo esperado (ambíguo)" for f in report.findings)


# ---------------------------------------------------------------------------
# Testes unitários — lista de tipos candidatos (coluna aceita mais de um tipo)
# ---------------------------------------------------------------------------
# Cenário real que motivou esta funcionalidade: "cliente_id" às vezes vem
# puramente numérico, às vezes como código alfanumérico ("CLI_109"),
# dependendo do export. tipos_esperados agora aceita uma lista para essas
# colunas, em vez de forçar um único tipo fixo.

def test_check_expected_types_lista_aceita_numerico(report):
    df = pd.DataFrame({"cliente_id": ["101", "102", "103"]})
    check_expected_types(df, {"cliente_id": ["numero", "texto"]}, report)

    assert not report.has_critical()


def test_check_expected_types_lista_aceita_alfanumerico(report):
    df = pd.DataFrame({"cliente_id": ["CLI_109", "CLI_136", "CLI_122"]})
    check_expected_types(df, {"cliente_id": ["numero", "texto"]}, report)

    # "texto" aceita qualquer valor, então nenhum dos dois formatos deve
    # ser marcado como crítico — a coluna bate com pelo menos um candidato.
    assert not report.has_critical()


def test_check_expected_types_lista_nenhum_candidato_bate(report):
    df = pd.DataFrame({"data_venda": ["2026-06-24", "2026-06-18", "2026-06-19"]})
    # "numero" e "data" propositalmente incompatíveis com o formato real
    # (strings de data já formatadas não batem com "numero").
    check_expected_types(df, {"data_venda": ["numero"]}, report)

    assert report.has_critical()


# ---------------------------------------------------------------------------
# Testes unitários — converter_tipos_colunas (tratamento) com lista de tipos
# ---------------------------------------------------------------------------

def test_converter_tipos_lista_usa_primeiro_candidato_compativel():
    df = pd.DataFrame({"cliente_id": ["101", "102", "103"]})
    resultado = converter_tipos_colunas(df.copy(), {"cliente_id": ["numero", "texto"]})

    assert resultado["cliente_id"].dtype != object
    assert pd.api.types.is_numeric_dtype(resultado["cliente_id"])


def test_converter_tipos_lista_cai_no_fallback_texto():
    """Caso que motivou a funcionalidade: cliente_id alfanumérico não deve
    ser destruído (virar tudo NaN) — deve cair no fallback 'texto' da lista
    e permanecer como estava, sem perda de dado."""
    df = pd.DataFrame({"cliente_id": ["CLI_109", "CLI_136", "CLI_122"]})

    with warnings.catch_warnings(record=True) as capturados:
        warnings.simplefilter("always")
        resultado = converter_tipos_colunas(df.copy(), {"cliente_id": ["numero", "texto"]})

    assert list(resultado["cliente_id"]) == ["CLI_109", "CLI_136", "CLI_122"]
    assert not resultado["cliente_id"].isna().any()
    # "texto" resolveu o candidato com sucesso; não deve gerar aviso de
    # "nenhum tipo compatível", já que o fallback funcionou.
    assert not any("IGNORADA" in str(w.message) for w in capturados)


def test_converter_tipos_sem_fallback_gera_aviso_e_preserva_coluna():
    """Regressão: sem 'texto' na lista de candidatos, se nada bate, a coluna
    deve ser preservada como estava (não virar tudo nulo) e um aviso deve
    ser emitido."""
    df = pd.DataFrame({"cliente_id": ["CLI_109", "CLI_136"]})

    with warnings.catch_warnings(record=True) as capturados:
        warnings.simplefilter("always")
        resultado = converter_tipos_colunas(df.copy(), {"cliente_id": ["numero"]})

    assert list(resultado["cliente_id"]) == ["CLI_109", "CLI_136"]
    assert any("IGNORADA" in str(w.message) for w in capturados)


# ---------------------------------------------------------------------------

def test_load_csv_utf8():
    buffer = io.BytesIO("coluna_a,coluna_b\n1,café".encode("utf-8-sig"))
    df = _load_csv(buffer)
    assert list(df.columns) == ["coluna_a", "coluna_b"]


def test_load_csv_fallback_latin1():
    """CORRIGIDO (auditoria): CSVs em latin-1 (comuns em exports brasileiros)
    devem ser lidos via fallback, não quebrar na primeira tentativa utf-8."""
    buffer = io.BytesIO("coluna_a,cidade\n1,São Paulo".encode("latin-1"))
    df = _load_csv(buffer)
    assert "São Paulo" in df["cidade"].values


def test_load_csv_reusa_buffer_entre_tentativas():
    """CORRIGIDO (auditoria): a função precisa dar seek(0) no buffer entre
    tentativas de encoding, senão a segunda tentativa lê a partir de onde a
    primeira parou (ou de nada) em vez do arquivo inteiro."""
    buffer = io.BytesIO("id,cidade\n1,Brasília".encode("latin-1"))
    df = _load_csv(buffer)
    assert len(df) == 1
    assert list(df.columns) == ["id", "cidade"]


def test_load_csv_ambos_encodings_falham_gera_erro_claro():
    """CORRIGIDO (auditoria): antes desta correção, esse caminho tentava
    `raise UnicodeDecodeError("mensagem")` com um único argumento, o que
    sempre disparava um TypeError (UnicodeDecodeError exige 5 argumentos
    posicionais) em vez de reportar o erro de fato."""
    with patch(
        "validador_planilhas.pd.read_csv",
        side_effect=UnicodeDecodeError("utf-8", b"\xff", 0, 1, "invalid"),
    ), pytest.raises(EncodingDetectionError):
        _load_csv("fake.csv")


# ---------------------------------------------------------------------------
# Testes unitários — deduplicação de nomes de aba
# ---------------------------------------------------------------------------

def test_resolve_duplicate_sheet_names_sem_colisao():
    resolved = resolve_duplicate_sheet_names(["Vendas", "Compras"])
    assert resolved == {"Vendas": "Vendas", "Compras": "Compras"}


def test_resolve_duplicate_sheet_names_case_insensitive():
    """ADICIONADO (sugestão de auditoria externa): 'Vendas' e 'vendas' não
    podem colidir silenciosamente — a segunda ocorrência ganha um sufixo."""
    resolved = resolve_duplicate_sheet_names(["Vendas", "vendas", "VENDAS"])
    assert resolved["Vendas"] == "Vendas"
    assert resolved["vendas"] == "vendas_1"
    assert resolved["VENDAS"] == "VENDAS_2"
    # Todas as chaves resultantes devem ser únicas.
    assert len(set(resolved.values())) == 3


def test_resolve_duplicate_sheet_names_trunca_31_caracteres():
    nome_longo = "Vendas_Regiao_Sudeste_2026_Completo"  # > 31 chars
    resolved = resolve_duplicate_sheet_names([nome_longo, nome_longo.lower()])
    for save_key in resolved.values():
        assert len(save_key) <= 31


def test_resolve_duplicate_sheet_names_sufixo_nao_viola_limite():
    """Caso de borda (item 11 da auditoria): nome com exatamente 30 caracteres
    + sufixo '_1' = 32 chars → deve ser truncado para 31. Garante que o
    resultado continua único (não cria nova colisão) e dentro do limite do Excel."""
    nome_30 = "A" * 30  # 30 chars exatos
    resolved = resolve_duplicate_sheet_names([nome_30, nome_30.lower()])
    save_keys = list(resolved.values())
    # Ambas as chaves dentro do limite do Excel
    assert all(len(k) <= 31 for k in save_keys)
    # As chaves devem ser únicas (sem sobrescrita)
    assert len(set(save_keys)) == 2


# ---------------------------------------------------------------------------
# Testes de integração — main() / CLI
# ---------------------------------------------------------------------------
# Estes testes cobrem especificamente os bugs que só aparecem no fluxo
# completo do script (parsing de argumentos, leitura de arquivo, decisão de
# formato de saída) — é exatamente onde estava o bug crítico original.

def test_cli_chave_duplicata_nao_quebra(tmp_path, monkeypatch):
    """Teste de regressão do bug crítico original: `args.chave-duplicata`
    era interpretado como subtração inválida e derrubava o script em toda
    execução que usasse --chave-duplicata."""
    arquivo = tmp_path / "vendas.xlsx"
    pd.DataFrame({"id_pedido": [1, 2, 2], "valor": [10, 20, 30]}).to_excel(arquivo, index=False)

    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, "argv", [
        "validador_planilhas.py", str(arquivo),
        "--chave-duplicata", "id_pedido",
        "--report", "rel.md",
    ])

    codigo = vp.main()

    # Duplicata real existe (id 2 repetido) -> deve retornar 1 (crítico),
    # nunca lançar exceção nem retornar 2 (erro de processamento).
    assert codigo == 1


def test_cli_sheet_indice_numerico(tmp_path, monkeypatch):
    """CORRIGIDO (auditoria): --sheet deve aceitar um índice numérico, não
    só o nome da aba."""
    arquivo = tmp_path / "multi.xlsx"
    with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
        pd.DataFrame({"a": [1]}).to_excel(writer, sheet_name="Primeira", index=False)
        pd.DataFrame({"b": [2, 3]}).to_excel(writer, sheet_name="Segunda", index=False)

    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, "argv", [
        "validador_planilhas.py", str(arquivo),
        "--sheet", "1",  # índice 1 = "Segunda"
        "--report", "rel.md",
    ])

    codigo = vp.main()

    assert codigo == 0
    conteudo = (tmp_path / "rel.md").read_text(encoding="utf-8")
    assert "Linhas: **2**" in conteudo  # a aba "Segunda" tem 2 linhas


def test_cli_all_sheets_dedup_case_insensitive(tmp_path, monkeypatch):
    """CORRIGIDO (sugestão de auditoria externa): abas cujos nomes só
    diferem em maiúsculas/minúsculas (ex.: 'Vendas' e 'vendas') não podem
    sobrescrever uma à outra silenciosamente em --all-sheets.

    Como o próprio Excel/openpyxl já impedem esse nome duplicado na
    escrita normal (confirmado manualmente), simulamos a leitura via mock
    para exercitar a lógica de deduplicação do script mesmo sem conseguir
    gerar um arquivo .xlsx real com essa colisão."""
    saida = tmp_path / "saida.xlsx"
    relatorio = tmp_path / "rel.md"

    fake_df = pd.DataFrame({"id_pedido": [1, 2]})

    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, "argv", [
        "validador_planilhas.py", "fake.xlsx",
        "--all-sheets", "--output", str(saida), "--report", str(relatorio),
    ])

    with patch("validador_planilhas.pd.ExcelFile", return_value=FakeExcelFile()), \
         patch("validador_planilhas.pd.read_excel", return_value=fake_df):
        codigo = vp.main()

    assert codigo == 0
    wb = openpyxl.load_workbook(saida)
    # As duas abas devem existir com nomes distintos, sem sobrescrita.
    assert wb.sheetnames == ["Vendas", "vendas_1"]


def test_cli_permission_error_mensagem_amigavel(tmp_path, monkeypatch, capsys):
    """ADICIONADO (sugestão de auditoria externa): PermissionError (arquivo
    aberto em outro programa) deve gerar mensagem amigável e código de saída
    2, não o erro genérico "Erro ao processar o arquivo: ..."."""
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, "argv", [
        "validador_planilhas.py", "bloqueado.xlsx", "--report", "rel.md",
    ])

    with patch("validador_planilhas.load_spreadsheet", side_effect=PermissionError("em uso")):
        codigo = vp.main()

    assert codigo == 2
    saida_erro = capsys.readouterr().err
    assert "aberto em outro programa" in saida_erro


def test_cli_output_csv_com_all_sheets_uma_aba(tmp_path, monkeypatch):
    """CORRIGIDO (auditoria): com --all-sheets, se o arquivo tiver apenas
    uma aba, salvar como .csv deve funcionar (antes dependia incorretamente
    da chave 'default', que nunca existe nesse fluxo)."""
    arquivo = tmp_path / "unica_aba.xlsx"
    pd.DataFrame({"id_pedido": [1, 2]}).to_excel(arquivo, index=False)
    saida_csv = tmp_path / "saida.csv"

    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, "argv", [
        "validador_planilhas.py", str(arquivo),
        "--all-sheets", "--output", str(saida_csv), "--report", "rel.md",
    ])

    codigo = vp.main()

    assert codigo == 0
    assert saida_csv.exists()
    assert pd.read_csv(saida_csv).shape[0] == 2


# ---------------------------------------------------------------------------
# Testes unitários — Segunda Auditoria (ações e fechamento de lacunas)
# ---------------------------------------------------------------------------

def test_render_markdown_report_com_renomeacao(report):
    """Ação 4: Testa render_markdown_report com column_rename_map preenchido."""
    report.column_rename_map = {"ID Pedido": "id_pedido"}
    report.add(SEVERITY_INFO, "Geral", "Tudo validado com sucesso.")
    md = render_markdown_report([report])
    assert "### Colunas renomeadas" in md
    assert "| ID Pedido | id_pedido |" in md


def test_check_mixed_types_ignora_coluna_booleana(report):
    """Ação 6: Colunas booleanas puras não devem disparar 'Tipos inconsistentes'."""
    df = pd.DataFrame({"ativo": [True, False, True]})
    check_mixed_types(df, report)
    assert not any(f.category == "Tipos inconsistentes" for f in report.findings)


def test_load_config_valido_e_invalido(tmp_path):
    """Ação 5: Testa load_config com None, arquivo válido e JSON inválido."""
    assert load_config(None) == {}

    cfg_file = tmp_path / "config.json"
    cfg_file.write_text('{"colunas_obrigatorias": ["id"]}', encoding="utf-8")
    assert load_config(str(cfg_file)) == {"colunas_obrigatorias": ["id"]}

    bad_cfg = tmp_path / "bad.json"
    bad_cfg.write_text("{json_invalido: 123", encoding="utf-8")
    with pytest.raises(SystemExit) as exc_info:
        load_config(str(bad_cfg))
    assert "Erro: config.json inválido" in str(exc_info.value)


def test_resolve_sheet_arg():
    """Ação 3: Converte índice numérico para int e mantém nome textual ou caracteres especiais."""
    assert _resolve_sheet_arg(None) is None
    assert _resolve_sheet_arg("1") == 1
    assert _resolve_sheet_arg("Vendas") == "Vendas"
    # Caractere unicode/superscript que falharia se fosse int() direto sem tratamento
    assert _resolve_sheet_arg("²") == "²"


def test_get_excel_sheets(tmp_path):
    """Ação 2: Lê as abas fechando o arquivo com context manager."""
    excel_path = tmp_path / "planilha_teste.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        pd.DataFrame({"a": [1]}).to_excel(writer, sheet_name="Aba1", index=False)
        pd.DataFrame({"b": [2]}).to_excel(writer, sheet_name="Aba2", index=False)
    sheets = get_excel_sheets(str(excel_path))
    assert sheets == ["Aba1", "Aba2"]


def test_check_nulls_df_vazio(report):
    """Ação 1: DataFrame vazio não itera nem gera achados de nulos."""
    df_vazio = pd.DataFrame()
    check_nulls(df_vazio, report)
    assert len(report.findings) == 0


def test_check_nulls_faixa_warning(report):
    """Lacuna de cobertura: nulos entre 5% e 30% devem gerar SEVERITY_WARNING."""
    df = pd.DataFrame({"col": [1, 2, 3, 4, 5, 6, 7, 8, 9, None]})  # 10% nulos
    check_nulls(df, report)
    warnings_found = [f for f in report.findings if f.severity == SEVERITY_WARNING]
    assert len(warnings_found) == 1
    assert "Coluna 'col' tem 10.0% de valores nulos." in warnings_found[0].message


def test_check_duplicate_rows_com_valores_nao_comparaveis(report):
    """Lacuna de cobertura: valores não comparáveis em célula disparam TypeError tratado."""
    df = pd.DataFrame({"dados": [1, 2]})
    with patch.object(df, "duplicated", side_effect=TypeError("tipo não comparável")):
        check_duplicate_rows(df, report)
    assert any("valores não comparáveis" in f.message for f in report.findings)


def test_check_duplicate_key_coluna_inexistente(report):
    """Lacuna de cobertura: checagem de chave quando a coluna não existe no DF."""
    df = pd.DataFrame({"id": [1, 2]})
    check_duplicate_key(df, "coluna_fantasma", report)
    assert any("não foi encontrada após padronização" in f.message for f in report.findings)


def test_check_mixed_types_coluna_numerica_pura_e_string_numerica(report):
    """Lacuna de cobertura: coluna numérica pura e coluna de strings numéricas puras."""
    df_num = pd.DataFrame({"numerica": [10, 20, 30]})
    check_mixed_types(df_num, report)
    assert len(report.findings) == 0

    df_str_num = pd.DataFrame({"str_num": ["10", "20", "30"]})
    check_mixed_types(df_str_num, report)
    assert len(report.findings) == 0


def test_compatibilidade_tipo_data_com_serie_numerica(report):
    """Lacuna de cobertura: _compatibilidade_tipo('data') com série numérica e string."""
    s_num = pd.Series([100, 200, 300])
    _compatibilidade_tipo(s_num, "data", "col_num", report)
    assert any("Coluna 'col_num' está configurada como tipo 'data'" in f.message for f in report.findings)

    report_str = ValidationReport(source_file="teste.xlsx")
    s_str = pd.Series(["100", "200", "300"])
    _compatibilidade_tipo(s_str, "data", "col_str", report_str)
    assert any("Coluna 'col_str' está configurada como tipo 'data'" in f.message for f in report_str.findings)


def test_check_expected_types_ausente_nula_e_sem_candidato(report):
    """Lacuna de cobertura: coluna ausente, coluna 100% nula e lista de tipos sem match."""
    df = pd.DataFrame({
        "col_nula": [None, None],
        "codigo": ["abc", "def"],
    })
    expected = {
        "col_inexistente": "numero",
        "col_nula": "numero",
        "codigo": ["numero", "data"],
    }
    check_expected_types(df, expected, report)
    # Apenas 'codigo' deve gerar aviso de tipo incompatível
    incompat_findings = [f for f in report.findings if "não é compatível com nenhum" in f.message]
    assert len(incompat_findings) == 1
    assert "codigo" in incompat_findings[0].message


def test_standardize_columns_colisao_gera_sufixo(report):
    """Lacuna de cobertura: duas colunas com nomes que colidem após normalização."""
    df = pd.DataFrame(columns=["Coluna A", "coluna_a"])
    df_padrao = standardize_columns(df, report)
    assert list(df_padrao.columns) == ["coluna_a", "coluna_a_1"]
    assert any("Coluna duplicada após padronização" in f.message for f in report.findings)


def test_standardize_column_name_vazio():
    """Lacuna de cobertura: nome de coluna vazio ou só espaços gera 'coluna_sem_nome'."""
    assert standardize_column_name("") == "coluna_sem_nome"
    assert standardize_column_name("   ") == "coluna_sem_nome"


def test_remover_colunas_vazias_sem_coluna_vazia():
    """Lacuna de cobertura: remover_colunas_vazias quando não há colunas vazias."""
    df = pd.DataFrame({"a": [1, 2], "b": ["x", "y"]})
    df_limpo = remover_colunas_vazias(df)
    assert df_limpo.shape == (2, 2)


def test_remover_duplicadas_chave_inexistente_emite_warning():
    """Ação 7: remover_duplicadas com chave inexistente emite UserWarning e dedup linhas."""
    df = pd.DataFrame({"a": [1, 1, 2], "b": [10, 10, 20]})
    with pytest.warns(UserWarning, match="não foi encontrada"):
        df_limpo = remover_duplicadas(df, chave="chave_fantasma")
    assert len(df_limpo) == 2


def test_tentar_converter_data_e_desconhecido():
    """Lacuna de cobertura: _tentar_converter para tipo 'data' e tipo desconhecido."""
    s = pd.Series(["21/09/2026", "22/09/2026"])
    conv_data = _tentar_converter(s, "data")
    assert conv_data is not None
    assert pd.api.types.is_datetime64_any_dtype(conv_data)
    assert _tentar_converter(s, "tipo_inexistente") is None


def test_converter_tipos_colunas_fallback_texto_e_desconhecido():
    """Lacuna de cobertura: converter_tipos_colunas com fallback 'texto' e tipo inválido."""
    df = pd.DataFrame({"col": [1, 2, 3]})
    df_conv = converter_tipos_colunas(df, {"col": "texto", "col_inex": "tipo_desconhecido"})
    assert "col" in df_conv.columns


def test_tratar_dataframe_completo_sem_tipos_esperados():
    """Lacuna de cobertura: tratar_dataframe_completo com tipos_esperados=None."""
    df = pd.DataFrame({"a": [1, 1, 2], "b": [None, None, None]})
    df_tratado = tratar_dataframe_completo(df, chave_duplicata=None, tipos_esperados=None)
    assert list(df_tratado.columns) == ["a"]
    assert len(df_tratado) == 2


def test_cli_arquivo_nao_encontrado(tmp_path, monkeypatch, capsys):
    """Lacuna de cobertura: CLI com arquivo inexistente."""
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, "argv", [
        "validador_planilhas.py", "arquivo_que_nao_existe.xlsx",
    ])
    codigo = vp.main()
    assert codigo == 2
    assert "arquivo não encontrado" in capsys.readouterr().err


def test_cli_erro_inesperado(tmp_path, monkeypatch, capsys):
    """Lacuna de cobertura: CLI com erro genérico inesperado."""
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, "argv", [
        "validador_planilhas.py", "qualquer.xlsx",
    ])
    with patch("validador_planilhas.load_spreadsheet", side_effect=RuntimeError("falha catastrófica")):
        codigo = vp.main()
    assert codigo == 2
    assert "falha catastrófica" in capsys.readouterr().err


def test_cli_all_sheets_com_output_csv_multiplas_abas_erro(tmp_path, monkeypatch, capsys):
    """Lacuna de cobertura: CLI com --all-sheets e --output .csv com múltiplas abas."""
    monkeypatch.chdir(tmp_path)
    fake_df = pd.DataFrame({"col": [1]})
    monkeypatch.setattr(sys, "argv", [
        "validador_planilhas.py", "fake.xlsx",
        "--all-sheets", "--output", "saida.csv",
    ])
    with patch("validador_planilhas.pd.ExcelFile", return_value=FakeExcelFile()), \
         patch("validador_planilhas.pd.read_excel", return_value=fake_df):
        codigo = vp.main()
    assert codigo == 2
    assert ".csv não suporta múltiplas abas" in capsys.readouterr().err


def test_cli_output_xlsx_com_tratamento(tmp_path, monkeypatch):
    """Lacuna de cobertura: CLI com --output .xlsx salvando dataframe tratado."""
    arquivo = tmp_path / "entrada.xlsx"
    saida = tmp_path / "saida.xlsx"
    pd.DataFrame({"id_pedido": [1, 1], "valor": [10, 10]}).to_excel(arquivo, index=False)
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, "argv", [
        "validador_planilhas.py", str(arquivo),
        "--output", str(saida),
    ])
    codigo = vp.main()
    assert codigo == 0
    assert saida.exists()
    assert pd.read_excel(saida).shape[0] == 1


def test_load_spreadsheet_csv_e_formato_invalido(tmp_path):
    csv_file = tmp_path / "dados.csv"
    csv_file.write_text("a,b\n1,2\n", encoding="utf-8")
    df = load_spreadsheet(str(csv_file))
    assert df.shape == (1, 2)

    with pytest.raises(ValueError, match="Formato não suportado"):
        load_spreadsheet("arquivo.parquet")


def test_check_mixed_types_edge_cases(report):
    # Coluna object vazia (somente NaN)
    df_empty_obj = pd.DataFrame({"vazia": [None, None]}, dtype=object)
    check_mixed_types(df_empty_obj, report)

    # Coluna object com bool, número, número como texto e texto puro misturados
    df_mixed = pd.DataFrame({"mista": [True, 123, "456", "texto"]}, dtype=object)
    check_mixed_types(df_mixed, report)
    assert any("Tipos inconsistentes" in f.category for f in report.findings)


def test_run_single_validation_com_expected_types():
    df = pd.DataFrame({"id": [1, 2], "nome": ["A", "B"]})
    df_val, rep = run_single_validation(df, source_name="teste.xlsx", expected_types={"id": "numero"})
    assert df_val.shape == (2, 2)
    assert not rep.has_critical()


def test_remover_duplicadas_chave_existente():
    df = pd.DataFrame({"id": [1, 1, 2], "val": [10, 20, 30]})
    df_dedup = remover_duplicadas(df, chave="id")
    assert len(df_dedup) == 2


def test_converter_tipos_colunas_tipo_desconhecido():
    df = pd.DataFrame({"col": [1, 2]})
    with pytest.warns(UserWarning, match="Conversão de 'col' foi IGNORADA"):
        df_res = converter_tipos_colunas(df, {"col": "desconhecido"})
    assert "col" in df_res.columns


def test_tratar_dataframe_completo_com_tipos():
    df = pd.DataFrame({"id": ["1", "2"], "data": ["01/01/2026", "02/01/2026"]})
    df_trat = tratar_dataframe_completo(df, chave_duplicata="id", tipos_esperados={"id": "numero", "data": "data"})
    assert df_trat.shape == (2, 2)