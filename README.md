# Validador e Padronizador de Planilhas para Power BI

Ferramenta em Python que valida, padroniza e **trata** planilhas (`.csv`, `.xlsx`, `.xls`) antes de entrarem no Power Query — pegando erros que normalmente só aparecem quando o relatório já está em produção.

Duas formas de uso:

- **CLI** (`validador_planilhas.py`) — para rodar em lote, scripts ou pipelines automatizados.
- **Interface web** (`app.py`, Streamlit) — para quem prefere subir o arquivo, ver o relatório na tela e baixar o resultado, sem linha de comando.

## Estrutura do projeto

```
.
├── validador_planilhas.py   # Motor de validação + CLI
├── tratador_planilhas.py    # Motor de tratamento (limpeza dos dados)
├── nomes_colunas.py         # Normalização de nomes de coluna (compartilhado pelos dois acima)
├── app.py                   # Interface web (Streamlit)
├── test_validador.py        # Suíte de testes (pytest)
├── config.exemplo.json      # Modelo de config.json
├── pyproject.toml           # Metadados, dependências, config de lint/teste
├── .gitignore
├── .streamlit/
│   └── config.toml          # Tema da interface web
└── imagens/                 # Imagens de fundo da interface (opcional)
```

`validador_planilhas.py` e `tratador_planilhas.py` são módulos separados por responsabilidade — um só **valida** (encontra problemas, gera relatório), o outro só **trata** (corrige os dados: remove coluna vazia, deduplica, converte tipo). `nomes_colunas.py` existe à parte porque os outros dois precisam da mesma normalização de nome de coluna, e um não pode importar do outro sem criar um import circular.

## O que ele verifica (validação)

- Nomes de colunas inconsistentes (espaços, acentos, maiúsculas, `\xa0`) → padroniza para `snake_case`
- Colunas obrigatórias ausentes (com match aproximado por similaridade)
- Duplicatas — linhas inteiras ou por uma coluna-chave
- Percentual de valores nulos acima do esperado
- Tipos de dados inconsistentes dentro da mesma coluna (texto misturado com número)
- Colunas 100% vazias (comum em exports mal feitos)
- Múltiplas abas de arquivos Excel, inclusive nomes de aba equivalentes ignorando maiúsculas/minúsculas

Gera um **relatório em Markdown** classificando cada achado como 🔴 Crítico, 🟡 Aviso ou 🔵 Info.

## O que ele corrige (tratamento)

Quando você passa `--output` (CLI) ou baixa a "Planilha Tratada" (interface web), os dados passam por `tratador_planilhas.py`:

- Remove colunas 100% vazias
- Remove linhas duplicadas (inteiras, ou por uma coluna-chave configurada)
- Converte colunas para o tipo esperado (`numero`, `data`), quando configurado

O relatório sempre reflete os dados **originais** — mostra os problemas encontrados antes do tratamento. O arquivo de saída é que sai efetivamente corrigido.

**Importante:** a conversão de tipo é defensiva. Se converter uma coluna para o tipo configurado apagaria mais da metade dos valores que eram válidos (sinal de que o tipo configurado não bate com os dados reais — ex.: `cliente_id` com valores tipo `"CLI_109"` configurado como `"numero"`), a conversão é **ignorada** e um aviso é emitido no terminal, em vez de destruir a coluna silenciosamente.

## Instalação

Instalação padrão (para rodar o script ou a interface):

```bash
pip install pandas openpyxl xlrd streamlit reportlab
```

Instalação como pacote editável (recomendado para desenvolvedores — expõe o comando `validador-planilhas` no terminal):

```bash
pip install -e .
```

Instalação com dependências de desenvolvimento (testes e linter):

```bash
pip install -e ".[dev]"
```

Requisitos: Python 3.10+ (o código usa `from __future__ import annotations` e sintaxe `str | None`).

## Uso — linha de comando (CLI)

Com o pacote instalado (`pip install -e .`):

```bash
validador-planilhas entrada.xlsx
```

Ou chamando o script diretamente, sem instalar:

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

> **Nota:** com `--all-sheets`, a saída (`--output`) só pode ser `.xlsx` se o arquivo tiver mais de uma aba (CSV não suporta múltiplas abas). Se o arquivo tiver apenas uma aba, `.csv` também funciona.

## Uso — interface web (Streamlit)

```bash
streamlit run app.py
```

Abre no navegador. Suba a planilha, opcionalmente um `config.json` na barra lateral, clique em "Executar Auditoria" e baixe o relatório (Markdown ou PDF) e a planilha tratada — no mesmo formato que você enviou (CSV → CSV, Excel → Excel).

O tema visual da interface (`.streamlit/config.toml`) e as imagens de fundo (pasta `imagens/`, opcional) ficam junto do `app.py`. Sem as imagens, a interface cai consideravelmente só no gradiente escuro, sem quebrar.

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
- **`tipos_esperados`**: dicionário `coluna → tipo`. Aceita um tipo único (`"numero"` ou `"data"`) ou uma **lista de tipos candidatos** (ex.: `["numero", "texto"]`), para colunas cujo formato varia entre exports diferentes da mesma planilha — a coluna é validada como ok se bater com qualquer um dos tipos da lista, e o tratamento tenta cada tipo na ordem dada, aplicando o primeiro que não perde dado demais. Incluir `"texto"` no fim da lista garante um fallback que nunca falha (não converte, preserva a coluna como está). Se menos de 95% dos valores forem compatíveis com **nenhum** dos tipos da lista, gera um achado crítico.
- **`chave_duplicata`**: alternativa a passar `--chave-duplicata` pela linha de comando (a opção de linha de comando tem prioridade). Também usada pelo tratamento para remover as duplicatas encontradas, não só reportá-las.

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

## Limitações conhecidas

- A checagem de tipos mistos (`check_mixed_types`) amostra até 50.000 linhas por coluna por questão de performance; em planilhas maiores, uma inconsistência muito rara pode não ser detectada.
- Datas são interpretadas com `dayfirst=True` (padrão brasileiro, `dd/mm/aaaa`); formatos americanos (`mm/dd/aaaa`) ambíguos podem ser interpretados incorretamente.
- O match aproximado de colunas obrigatórias é baseado em similaridade textual (`difflib`) e pode, em casos raros, casar com a coluna errada se os nomes forem muito parecidos — sempre revise o relatório antes de confiar no rename automático.
- A conversão de tipo é ignorada (não aplicada) quando causaria perda de mais de 50% dos valores válidos de uma coluna — isso evita destruir dados quando o tipo configurado está errado, mas significa que **você** ainda precisa corrigir o `config.json` para aquela coluna ser tratada de fato; o script não adivinha o tipo certo sozinho.
- Colunas com valores realmente ausentes na origem (célula vazia) ou com texto que não corresponde a nenhum tipo reconhecido (ex.: um valor numérico escrito por extenso) continuam nulas após o tratamento — "sem nulos" é uma decisão de negócio (remover a linha, preencher com um valor padrão, ou aceitar o gap), não algo que o script resolve sozinho.
- `st.dataframe` (usado para exibir o mapeamento de colunas renomeadas na interface web) renderiza num componente próprio do Streamlit que não herda CSS customizado — o tema escuro dessa tabela específica depende do `.streamlit/config.toml`, não do CSS do `app.py`.

## Changelog

### Interface web e tratamento de dados
- Adicionada interface Streamlit (`app.py`) com upload, relatório na tela, veredito visual, download do relatório (Markdown/PDF) e da planilha tratada.
- Adicionado `tratador_planilhas.py`: remoção de coluna vazia, deduplicação por chave, conversão de tipo — antes só a validação existia, sem correção automática dos dados.
- `nomes_colunas.py` extraído como módulo compartilhado (sem dependências) para `validador_planilhas.py` e `tratador_planilhas.py` usarem a mesma normalização de nome de coluna sem criar import circular entre eles.
- Corrigido bug crítico: uma versão anterior de `tratador_planilhas.py` normalizava nomes de coluna de forma mais fraca que `validador_planilhas.py` (não removia acento, não trocava espaço por `_`), fazendo a remoção de duplicata e a conversão de tipo falharem silenciosamente sempre que o `config.json` tinha um nome de coluna com espaço ou acento.
- Corrigido bug crítico: conversão de tipo aplicada sem checagem podia zerar uma coluna inteira (ex.: `cliente_id` no formato `"CLI_109"` configurado como `"numero"` virava 100% nulo). Agora a conversão só é aplicada se não perder mais de 50% dos valores válidos; senão, é ignorada e um aviso é emitido.
- Adicionado suporte a lista de tipos candidatos em `tipos_esperados` (ex.: `["numero", "texto"]`), para colunas cujo formato varia entre exports diferentes.
- Corrigida perda de estado na interface web: resultados desapareciam da tela ao clicar em qualquer botão de download (comportamento do Streamlit — qualquer interação dispara um rerun completo do script). Resultados agora vivem em `st.session_state`.
- Corrigido: a interface web validava e renomeava colunas, mas nunca aplicava o tratamento antes de gerar o download — o arquivo "tratado" saía sem tratamento nenhum.
- Corrigida leitura de CSV na interface web: reimplementava `pd.read_csv` sem o fallback de encoding (utf-8-sig → latin-1) já usado no CLI, quebrando em exports latin-1 comuns no Brasil.
- Corrigida deduplicação de abas com nomes equivalentes ignorando maiúsculas/minúsculas (ex.: `"Vendas"`/`"vendas"`) também na interface web, reaproveitando `resolve_duplicate_sheet_names` do CLI em vez de duplicar a lógica.

### Empacotamento e infraestrutura
- Corrigido `pyproject.toml`: faltava declarar `xlrd`, a biblioteca que o pandas exige para ler `.xls` (formato antigo do Excel — `openpyxl` só lê `.xlsx`). O CLI e o README anunciavam suporte a `.xls`, mas sem essa dependência o script quebrava com `ImportError` ao processar um `.xls` real. Confirmado gerando e testando um arquivo `.xls` de verdade.
- Corrigido `pyproject.toml`: `py-modules` listava só `validador_planilhas`, mas o projeto tem três módulos interdependentes — `pip install .` instalava um pacote incompleto que quebrava com `ModuleNotFoundError` ao ser importado de fora do diretório do projeto.
- Adicionado `.gitignore` cobrindo ambientes virtuais (`.venv/`, `venv/`), `.streamlit/secrets.toml`, arquivos de sistema operacional, e os nomes reais de saída gerados pelo CLI e pela interface web.
- Adicionado `.streamlit/config.toml` com tema escuro, para que componentes nativos do Streamlit que não herdam CSS customizado (como `st.dataframe`) fiquem consistentes com o resto da interface.

### Revisão de auditoria (motor de validação)
- Corrigido bug crítico que impedia a execução do script (`--chave-duplicata` sempre causava erro).
- `--sheet` agora aceita índice numérico, não só nome da aba.
- Corrigida a decisão de formato de saída ao usar `--all-sheets` com uma única aba.
- `pd.ExcelFile` é aberto uma única vez ao validar múltiplas abas (evita releituras).
- Match aproximado de coluna obrigatória agora renomeia a coluna de fato, não só avisa.
- Checagem de tipos mistos agora varre a coluna inteira (até um teto de performance), não só as primeiras 500 linhas.
- Removida duplicação de achados entre coluna 100% vazia e coluna 100% nula.
- Leitura de CSV não mascara mais erros não relacionados a encoding, e usa uma exceção própria (`EncodingDetectionError`) em vez de uma construção inválida de `UnicodeDecodeError`.
- Abas de Excel com nomes equivalentes ignorando maiúsculas/minúsculas não sobrescrevem mais uma à outra silenciosamente em `--all-sheets`.
- `check_expected_types` sinaliza quando uma coluna marcada como tipo `"data"` tem uma fração relevante de valores puramente numéricos (ex.: número de série do Excel não convertido).
- `PermissionError` (arquivo aberto em outro programa) e `FileNotFoundError` geram mensagens amigáveis e específicas no terminal, em vez de caírem no erro genérico.