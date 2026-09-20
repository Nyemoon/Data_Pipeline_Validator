"""
Módulo de Tratamento e Correção de Erros de Planilhas
=====================================================
Funções utilitárias puras para tratar e limpar dataframes do pandas.
"""

from __future__ import annotations
import pandas as pd
import warnings

from nomes_colunas import standardize_column_name

# Se a conversão de tipo transformaria em nulo mais do que esta fração dos
# valores que eram válidos antes (não-nulos), a coluna inteira é preservada
# como estava, em vez de aplicar a conversão. Um punhado de valores ruins
# virando nulo é limpeza normal; a maioria (ou tudo) virando nulo é sinal de
# que o tipo configurado não corresponde aos dados reais da coluna.
LIMITE_PERDA_ACEITAVEL = 0.5


def remover_colunas_vazias(df: pd.DataFrame) -> pd.DataFrame:
    """Remove colunas que estão 100% vazias (NaN)."""
    colunas_vazias = [c for c in df.columns if df[c].isna().all()]
    if colunas_vazias:
        df = df.drop(columns=colunas_vazias)
    return df

def remover_duplicadas(df: pd.DataFrame, chave: str | None = None) -> pd.DataFrame:
    """Remove linhas inteiramente duplicadas ou duplicadas com base em uma chave.

    CORRIGIDO: `chave` agora passa pela mesma normalização usada para
    renomear as colunas do DataFrame (standardize_column_name). Antes, uma
    chave como "ID Pedido" (com espaço, do jeito que veio do config.json)
    nunca batia com a coluna real já renomeada para "id_pedido", e a função
    caía silenciosamente no dedup por linha inteira em vez de por chave.
    """
    if chave:
        chave_normalizada = standardize_column_name(chave)
        if chave_normalizada in df.columns:
            return df.drop_duplicates(subset=[chave_normalizada], keep="first")
    return df.drop_duplicates()

def _tentar_converter(series: pd.Series, tipo: str) -> pd.Series | None:
    """Tenta converter `series` para `tipo`. Retorna a série convertida, ou
    None se `tipo` não é um tipo reconhecido para conversão. "texto" nunca
    converte (retorna a série original) — serve como fallback seguro dentro
    de uma lista de tipos candidatos, já que nunca perde dado."""
    if tipo == "numero":
        return pd.to_numeric(series, errors="coerce")
    if tipo == "data":
        return pd.to_datetime(series, errors="coerce", dayfirst=True)
    if tipo == "texto":
        return series
    return None


def converter_tipos_colunas(
    df: pd.DataFrame, tipos_esperados: dict[str, str | list[str]]
) -> pd.DataFrame:
    """Converte colunas para os tipos esperados (ex: 'numero', 'data').

    CORRIGIDO: usa standardize_column_name (a mesma função usada para
    renomear as colunas do DataFrame), não apenas `.strip().lower()`. A
    versão anterior não removia acentos nem trocava espaço por "_", então
    uma chave de config como "Data Venda" ou "Preço Unitário" nunca batia
    com a coluna já padronizada ("data_venda", "preco_unitario") e a
    conversão de tipo era silenciosamente pulada.

    CORRIGIDO (2ª rodada — confirmado com dados reais): antes desta
    correção, a conversão era aplicada cegamente. Uma coluna como
    "cliente_id" com valores tipo "CLI_109" configurada como "numero" virava
    100% NaN — a coluna inteira era destruída, silenciosamente, sem aviso
    nenhum. Agora a função mede quantos valores que ERAM válidos (não-nulos)
    virariam nulo com a conversão; se a perda passar de
    LIMITE_PERDA_ACEITAVEL, o tipo é descartado (e o próximo candidato,
    se houver, é tentado) em vez de aplicar uma conversão que apaga a coluna.

    ADICIONADO: cada coluna em `tipos_esperados` agora pode apontar para uma
    lista de tipos candidatos (ex.: ["numero", "texto"]), não só um único
    tipo — útil para colunas cujo formato varia entre exports diferentes da
    mesma planilha (ex.: cliente_id às vezes puramente numérico, às vezes um
    código alfanumérico tipo "CLI_109"). Os candidatos são tentados na ordem
    dada; o primeiro que não perde dado demais é aplicado, e os demais são
    ignorados. Incluir "texto" no fim da lista garante um fallback que nunca
    falha (não converte, preserva a coluna como está).
    """
    for raw_col, tipo_ou_lista in tipos_esperados.items():
        col = standardize_column_name(raw_col)
        if col not in df.columns:
            continue

        candidatos = tipo_ou_lista if isinstance(tipo_ou_lista, list) else [tipo_ou_lista]
        original = df[col]
        validos_antes = original.notna().sum()
        aplicado = False

        for tipo in candidatos:
            convertido = _tentar_converter(original, tipo)
            if convertido is None:
                continue  # tipo não reconhecido para conversão; pula

            if tipo == "texto":
                # Nunca converte, nunca perde dado — fallback sempre válido.
                aplicado = True
                break

            validos_depois = convertido.notna().sum()
            perda = 1 - (validos_depois / validos_antes) if validos_antes > 0 else 0.0

            if perda <= LIMITE_PERDA_ACEITAVEL:
                df[col] = convertido
                aplicado = True
                break
            # Perda alta demais: descarta este candidato e tenta o próximo.

        if not aplicado:
            tentativas = ", ".join(candidatos)
            warnings.warn(
                f"Conversão de '{col}' foi IGNORADA: nenhum dos tipos "
                f"candidatos ({tentativas}) é compatível o suficiente com os "
                f"dados reais da coluna (ex.: "
                f"'{original.dropna().iloc[0] if validos_antes else ''}'). "
                f"Revise o tipos_esperados no config.json.",
                stacklevel=2,
            )

    return df

def tratar_dataframe_completo(df: pd.DataFrame, chave_duplicata: str | None = None, tipos_esperados: dict[str, str] | None = None) -> pd.DataFrame:
    """Aplica todas as rotinas de tratamento padrão de forma encadeada."""
    df = remover_colunas_vazias(df)
    df = remover_duplicadas(df, chave=chave_duplicata)
    if tipos_esperados:
        df = converter_tipos_colunas(df, tipos_esperados)
    return df