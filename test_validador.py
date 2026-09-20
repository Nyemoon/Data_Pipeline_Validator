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
import json
import sys
import warnings
from unittest.mock import patch

import openpyxl
import pandas as pd
import pytest

import validador_planilhas as vp
from validador_planilhas import (
    SEVERITY_CRITICAL,
    EncodingDetectionError,
    ValidationReport,
    _load_csv,
    check_duplicate_key,
    check_duplicate_rows,
    check_empty_columns,
    check_expected_types,
    check_mixed_types,
    check_nulls,
    check_required_columns,
    resolve_duplicate_sheet_names,
    standardize_columns,
)
from tratador_planilhas import converter_tipos_colunas


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
    ):
        with pytest.raises(EncodingDetectionError):
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

    class FakeExcelFile:
        sheet_names = ["Vendas", "vendas"]

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