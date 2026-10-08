"""
Normalização de nomes de coluna — compartilhada entre validador_planilhas.py
e tratador_planilhas.py
==============================================================================

Extraído para seu próprio módulo, sem depender de nenhum dos outros dois,
justamente para que ambos possam importar a MESMA função de normalização
sem criar uma importação circular (validador -> tratador -> validador).

Antes desta extração, tratador_planilhas.py reimplementava a normalização
de forma mais fraca (`raw_col.strip().lower()`, sem remover acentos nem
trocar espaço por "_"), o que fazia colunas do config como "Data Venda" ou
"ID Pedido" nunca baterem com o nome já padronizado no DataFrame
("data_venda", "id_pedido") — e o tratamento (conversão de tipo, remoção de
duplicata por chave) falhava silenciosamente, sem erro nenhum.
"""

from __future__ import annotations

import re
import unicodedata


def remove_accents(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(c for c in normalized if not unicodedata.combining(c))


def standardize_column_name(name: str) -> str:
    # Substitui espaços não-quebráveis (\xa0) e outros whitespaces por espaço comum
    name = str(name).replace("\xa0", " ")
    name = name.strip()
    name = remove_accents(name)
    name = name.lower()
    name = re.sub(r"[^\w\s]", "_", name)  # pontuação -> underscore
    name = re.sub(r"\s+", "_", name)  # espaços -> underscore
    name = re.sub(r"_+", "_", name)  # underscores repetidos
    name = name.strip("_")
    if not name:
        name = "coluna_sem_nome"
    return name
