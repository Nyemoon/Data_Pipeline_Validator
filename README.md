# Validador e Padronizador de Planilhas para Power BI

Ferramenta em Python que valida, padroniza e **trata** planilhas (`.csv`, `.xlsx`, `.xls`) antes de entrarem no Power Query — pegando erros que normalmente só aparecem quando o relatório já está em produção.

Duas formas de uso:

- **CLI** (`validador_planilhas.py`) — para rodar em lote, scripts ou pipelines automatizados.
- **Interface web** (`app.py`, Streamlit) — para quem prefere subir o arquivo, ver o relatório na tela e baixar o resultado, sem linha de comando.

---

## Estrutura do projeto

```
.
├── validador_planilhas.py       # Motor de validação + CLI
├── tratador_planilhas.py        # Motor de tratamento (limpeza dos dados)
├── nomes_colunas.py             # Normalização de nomes de coluna (compartilhado pelos dois acima)
├── app.py                       # Interface web (Streamlit)
├── test_validador.py            # Suíte de testes (pytest: 62 testes, 99.8% de cobertura)
├── config.exemplo.json          # Modelo de config.json
├── pyproject.toml               # Metadados, dependências, config de lint/teste
├── .github/
│   └── workflows/
│       └── validate_sheets.yml  # Pipeline de CI (testes automatizados + ruff)
├── .gitignore
├── .streamlit/
│   └── config.toml              # Tema da interface web
└── imagens/                     # Imagens de fundo da interface (opcional)
```

`validador_planilhas.py` e `tratador_planilhas.py` são módulos separados por responsabilidade — um só **valida** (encontra problemas, gera relatório), o outro só **trata** (corrige os dados: remove coluna vazia, deduplica, converte tipo). `nomes_colunas.py` existe à parte porque os outros dois precisam da mesma normalização de nome de coluna, e um não pode importar do outro sem criar um import circular.

---

## O que ele verifica (validação)

- **Nomes de colunas inconsistentes** (espaços extras, acentos, maiúsculas/minúsculas, caracteres invisíveis como `\xa0`) → padroniza para `snake_case`.
- **Colunas obrigatórias ausentes** (com match aproximado por similaridade textual `difflib` ≥ 80% e renomeação automática).
- **Duplicatas** — linhas inteiramente duplicadas ou por uma coluna-chave configurada.
- **Percentual de valores nulos** — alerta avisos (≥ 5%) e problemas críticos (≥ 30%), com tratamento para DataFrames vazios.
- **Tipos de dados inconsistentes** dentro da mesma coluna (texto misturado com número, booleano ou números armazenados como texto).
- **Colunas 100% vazias** (comuns em exports mal formatados).
- **Múltiplas abas de arquivos Excel**, inclusive nomes de aba equivalentes ignorando maiúsculas/minúsculas.
- **Codificação de arquivos CSV**: detecção com triplo fallback (`utf-8-sig` → `latin-1` → `cp1252`).

Gera um **relatório em Markdown** classificando cada achado como 🔴 Crítico, 🟡 Aviso ou 🔵 Info.

---

## O que ele corrige (tratamento)

Quando você passa `--output` (CLI) ou baixa a "Planilha Tratada" (interface web), os dados passam por `tratador_planilhas.py`:

- Remove colunas 100% vazias.
- Remove linhas duplicadas (inteiras, ou pela coluna-chave configurada). Se a coluna-chave não existir após a padronização, um aviso (`UserWarning`) é emitido e o fallback de linha inteira é aplicado.
- Converte colunas para o tipo esperado (`numero`, `data`), quando configurado.

O relatório sempre reflete os dados **originais** — mostra os problemas encontrados antes do tratamento. O arquivo de saída é que sai efetivamente corrigido.

> [!IMPORTANT]
> A conversão de tipo é **defensiva**: se converter uma coluna para o tipo configurado apagaria mais de 50% dos valores que eram válidos antes (sinal de que o tipo configurado não bate com os dados reais — ex.: `cliente_id` com valores tipo `"CLI_109"` configurado como `"numero"`), a conversão é **ignorada** e um aviso explicativo é emitido no terminal ou logs, em vez de destruir a coluna silenciosamente.

---

## Instalação e Requisitos

Requisitos: **Python 3.10+** (usa `from __future__ import annotations` e anotações modernas `str | None`).

### Instalação padrão (para rodar o script ou a interface)

```bash
pip install pandas openpyxl xlrd streamlit reportlab
```

### Instalação como pacote editável (recomendado para desenvolvedores)

Expõe o executável `validador-planilhas` no PATH:

```bash
pip install -e .
```

### Instalação com dependências de desenvolvimento (testes e linter)

```bash
pip install -e ".[dev]"
```

---

## Testes, Cobertura e Qualidade de Código

O projeto conta com uma suíte abrangente de **62 testes unitários e de integração** com **99.8% de cobertura**:

```bash
# Rodar todos os testes
pytest test_validador.py -v

# Rodar testes com relatório de cobertura de código
pytest test_validador.py --cov=validador_planilhas --cov=tratador_planilhas --cov=nomes_colunas --cov-report=term-missing

# Verificar conformidade de estilo e boas práticas com ruff
ruff check .
```

O repositório inclui automação de Integração Contínua (CI) via GitHub Actions (`.github/workflows/validate_sheets.yml`), validando automaticamente testes e linting em cada pull request e push.

---

## Uso — linha de comando (CLI)

Com o pacote instalado (`pip install -e .`):

```bash
validador-planilhas entrada.xlsx
```

Ou chamando o script diretamente:

```bash
python validador_planilhas.py entrada.xlsx
```

Isso valida a primeira aba do arquivo e grava o relatório em `relatorio.md` (padrão).

### Opções

| Opção | Descrição |
|---|---|
| `entrada` | Caminho do arquivo `.csv`, `.xlsx` ou `.xls` (obrigatório) |
| `--config CONFIG` | Caminho de um `config.json` com colunas obrigatórias e tipos esperados |
| `--output OUTPUT` | Caminho para salvar a planilha **tratada** (`.xlsx` ou `.csv`) |
| `--report REPORT` | Caminho do relatório em Markdown (padrão: `relatorio.md`) |
| `--chave-duplicata COLUNA` | Nome de uma coluna a checar como chave única (detecta e remove duplicatas por valor) |
| `--sheet SHEET` | Nome **ou índice numérico** da aba do Excel a validar (ignorado para `.csv`) |
| `--all-sheets` | Valida todas as abas do arquivo Excel, uma a uma |

### Exemplos

```bash
# Validar uma aba específica pelo nome
python validador_planilhas.py entrada.xlsx --sheet "Vendas_2026"

# Validar uma aba pelo índice (0 = primeira aba)
python validador_planilhas.py entrada.xlsx --sheet 1

# Validar todas as abas e salvar tudo já tratado
python validador_planilhas.py entrada.xlsx --all-sheets --output limpo.xlsx

# Usar um config.json com regras de validação
python validador_planilhas.py entrada.csv --config config.json

# Checar e remover duplicatas por uma coluna-chave
python validador_planilhas.py entrada.xlsx --chave-duplicata id_pedido --output limpo.xlsx
```

> [!NOTE]
> Com `--all-sheets`, a saída (`--output`) deve ser `.xlsx` caso o arquivo contenha múltiplas abas (CSV não suporta múltiplas abas). Se o arquivo contiver apenas uma aba, `.csv` é aceito normalmente.

---

## Uso — interface web (Streamlit)

```bash
streamlit run app.py
```

Abre no navegador. Suba a planilha, opcionalmente um `config.json` na barra lateral, clique em "Executar Auditoria" e baixe o relatório (Markdown ou PDF) e a planilha tratada — mantendo o formato correspondente (CSV → CSV, Excel → Excel).

O tema visual da interface (`.streamlit/config.toml`) e as imagens de fundo (pasta `imagens/`, opcional) ficam junto do `app.py`.

---

## Formato do `config.json`

```json
{
  "colunas_obrigatorias": [
    "id_pedido",
    "data_venda",
    "valor_total",
    "cliente_id"
  ],
  "tipos_esperados": {
    "id_pedido": "numero",
    "data_venda": "data",
    "valor_total": "numero",
    "cliente_id": ["numero", "texto"]
  },
  "chave_duplicata": "id_pedido"
}
```

- **`colunas_obrigatorias`**: lista de colunas que devem existir. O script tenta casar por nome padronizado e, se não achar exato, por similaridade ≥ 80%, renomeando a coluna encontrada.
- **`tipos_esperados`**: dicionário `coluna → tipo`. Aceita um tipo único (`"numero"` ou `"data"`) ou uma **lista de tipos candidatos** (ex.: `["numero", "texto"]`), para colunas cujo formato varia entre exports diferentes da mesma planilha — a coluna é validada como ok se bater com qualquer um dos tipos da lista, e o tratamento tenta cada tipo na ordem dada, aplicando o primeiro que não perde dados. Incluir `"texto"` no fim da lista garante um fallback que nunca falha (não converte, preserva a coluna como está).
- **`chave_duplicata`**: alternativa a passar `--chave-duplicata` pela linha de comando (a opção de linha de comando tem prioridade). Também usada pelo tratamento para remover as duplicatas encontradas.

---

## Códigos de saída (CLI)

| Código | Significado |
|---|---|
| `0` | Nenhum problema crítico encontrado |
| `1` | Pelo menos um problema crítico foi encontrado em alguma aba/arquivo |
| `2` | Erro ao processar o arquivo (arquivo inválido, corrompido, encoding não detectado, permissão negada, config malformado etc.) |

Isso permite usar o script como *gate* em pipelines automatizados:

```bash
python validador_planilhas.py entrada.xlsx || echo "Corrija os problemas antes de publicar no Power BI"
```

---

## Limitações conhecidas

- A checagem de tipos mistos (`check_mixed_types`) amostra até 50.000 linhas por coluna por questão de performance; em planilhas maiores, uma inconsistência muito rara pode não ser detectada.
- Datas são interpretadas com `dayfirst=True` (padrão brasileiro, `dd/mm/aaaa`); formatos americanos (`mm/dd/aaaa`) ambíguos podem ser interpretados incorretamente se não especificados no esquema.
- O match aproximado de colunas obrigatórias é baseado em similaridade textual (`difflib`) e pode, em casos raros, casar com a coluna errada se os nomes forem muito parecidos — sempre revise o relatório antes de confiar no rename automático.
- A conversão de tipo é ignorada (não aplicada) quando causaria perda de mais de 50% dos valores válidos de uma coluna — isso evita destruir dados quando o tipo configurado está errado, mas significa que o esquema precisa de ajuste no `config.json`.
- Colunas com valores ausentes na origem continuam nulas após o tratamento — preenchimento de nulos depende de regra de negócio específica.

---

## Changelog

### Auditorias e Refinamento Contínuo (Qualidade, Robustez e Cobertura)
- **Cobertura de testes de 99.8%**: expansão da suíte para 62 testes automatizados cobrindo todos os fluxos de validação, tratamento, CLI e edge cases.
- **Integração Contínua (CI)**: adicionado workflow do GitHub Actions com validação de testes (`pytest`) e análise estática (`ruff check .`).
- **Triplo fallback de encoding em CSV**: adicionado fallback para `cp1252` (Windows-1252) após `utf-8-sig` e `latin-1`, garantindo decodificação correta de arquivos gerados por versões antigas do Excel no Windows.
- **Fechamento seguro de arquivos Excel**: `get_excel_sheets` utiliza context manager `with pd.ExcelFile` garantindo que descritores de arquivos não fiquem abertos (prevenindo vazamento de recursos e bloqueio em sistemas Windows).
- **Tratamento gracioso de JSONDecodeError**: tanto no CLI (`load_config`) quanto na interface web (`app.py`), arquivos JSON malformados geram mensagens de erro claras e amigáveis sem expor tracebacks não tratados.
- **Aviso explícito em deduplicação por chave**: `remover_duplicadas` agora emite um `UserWarning` quando a coluna informada como chave única não é encontrada após a padronização, evitando fallback silencioso para dedup de linha inteira.
- **Resolução robusta de argumentos de aba**: `_resolve_sheet_arg` usa conversão com `try/except ValueError`, suportando numerais Unicode e caracteres especiais sem quebrar.
- **Tratamento de DataFrames vazios**: otimizada a verificação em `check_nulls` para retornar imediatamente caso o DataFrame tenha 0 linhas e/ou colunas.
- **Silenciamento controlado de warnings de data**: `tratador_planilhas.py` silencia de forma segura avisos inofensivos de parsing em formato `dayfirst=True`.

### Interface web e tratamento de dados
- Adicionada interface Streamlit (`app.py`) com upload, relatório na tela, veredito visual, download do relatório (Markdown/PDF) e da planilha tratada.
- Adicionado `tratador_planilhas.py`: remoção de coluna vazia, deduplicação por chave, conversão de tipo.
- `nomes_colunas.py` extraído como módulo compartilhado (sem dependências) para `validador_planilhas.py` e `tratador_planilhas.py` usarem a mesma normalização de nome de coluna sem criar import circular entre eles.
- Corrigido bug crítico: normalização de nomes de coluna em `tratador_planilhas.py` alinhada à do validador (removendo acentos e espaços), eliminando falhas silenciosas de deduplicação por chave.
- Corrigido bug crítico: conversão de tipo com proteção contra perda excessiva de dados (limite aceitável de 50%).
- Adicionado suporte a lista de tipos candidatos em `tipos_esperados` (ex.: `["numero", "texto"]`).
- Corrigida persistência de estado na interface web utilizando `st.session_state` para evitar perda de dados em reruns do Streamlit.
- Adicionado tratamento real no fluxo de download da interface web.
- Unificação do algoritmo de resolução de nomes de aba duplicados (`resolve_duplicate_sheet_names`) entre CLI e web.

### Empacotamento e infraestrutura
- Declarada dependência `xlrd` no `pyproject.toml` para permitir leitura de arquivos `.xls` legado.
- Corrigida declaração de `py-modules` no `pyproject.toml` para incluir todos os módulos interdependentes do pacote.
- Adicionado `.gitignore` cobrindo ambientes virtuais (`.venv/`, `venv/`), `.streamlit/secrets.toml`, e arquivos de saída.
- Adicionado `.streamlit/config.toml` com tema escuro para consistência visual em todos os componentes nativos do Streamlit.

### Motor de validação (CLI)
- Corrigido bug em `--chave-duplicata` que causava erro de sintaxe em tempo de execução.
- `--sheet` aceita índice numérico ou nome de aba.
- Abertura única de `pd.ExcelFile` ao processar múltiplas abas (`--all-sheets`).
- Match aproximado de colunas obrigatórias com renomeação efetiva no DataFrame.
- Varredura completa para detecção de tipos mistos em colunas do tipo objeto.
- Eliminação de duplicidade de achados entre colunas 100% vazias e nulas.
- Mensagens amigáveis no terminal para `PermissionError` (arquivo em uso) e `FileNotFoundError`.