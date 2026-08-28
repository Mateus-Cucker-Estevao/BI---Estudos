# 📊 Estudos em Power BI

Repositório pessoal para registrar aprendizados, códigos e boas práticas de **Power BI**, **Power Query (M)** e **DAX**.

Cada tópico novo que eu aprendo vira uma seção aqui, com explicação, exemplo prático e o motivo de usar.

---

## 📁 Estrutura do repositório

```
/
├── README.md              # Este arquivo (índice geral dos estudos)
├── /power-query           # Códigos M, funções personalizadas e transformações
├── /dax                   # Medidas, colunas calculadas e tabelas calculadas
├── /modelagem             # Diagramas, relacionamentos e boas práticas de modelo
├── /pbix                  # Arquivos .pbix de exercícios e projetos
└── /assets                # Prints, imagens e materiais de apoio
```

---

## 🗂️ Índice de aprendizados

| # | Tópico | Área | Status |
|---|--------|------|--------|
| 01 | [Parâmetros: centralizando o caminho das pastas](#01--parâmetros-centralizando-o-caminho-das-pastas) | Power Query | ✅ |
| 02 | [Tabela calendário (dim_calendario) com parâmetros de data](#02--tabela-calendário-dim_calendario-com-parâmetros-de-data) | Power Query | ✅ |

---

## 01 — Parâmetros: centralizando o caminho das pastas

### O problema
Quando cada consulta aponta para um caminho fixo (`C:\Users\SeuNome\Documentos\Base\vendas.xlsx`), qualquer mudança de pasta ou troca de computador quebra **todas** as consultas de uma vez. Aí é corrigir uma por uma na mão.

### A solução
Criar um **parâmetro** com o caminho base e usar ele como referência em todas as consultas. Muda o caminho em **um lugar só** e o modelo inteiro se ajusta.

### Como fazer

1. Abrir o **Power Query** (Transformar dados)
2. Guia **Página Inicial** → **Gerenciar Parâmetros** → **Novo Parâmetro**
3. Preencher:
   - **Nome:** `CaminhoBase` (sem espaços e sem acentos — facilita usar no código M)
   - **Tipo:** `Texto`
   - **Valores sugeridos:** `Qualquer valor` (ou `Lista de valores` se quiser alternar entre pastas)
   - **Valor atual:** `C:\Users\SeuNome\Documentos\Base\`
4. Ir na consulta → **Editor Avançado** → substituir o caminho fixo pelo parâmetro

### Exemplo em código M

**Antes (caminho fixo — evitar):**
```m
let
    Fonte = Excel.Workbook(File.Contents("C:\Users\SeuNome\Documentos\Base\vendas.xlsx"), null, true)
in
    Fonte
```

**Depois (usando parâmetro):**
```m
let
    Fonte = Excel.Workbook(File.Contents(CaminhoBase & "vendas.xlsx"), null, true)
in
    Fonte
```

> O operador `&` concatena textos no M. Atenção à barra `\` no fim do parâmetro: ou ela está no parâmetro, ou no nome do arquivo — nunca nas duas pontas (senão vira `\\`).

### Exemplo com pasta inteira (Folder.Files)

Muito usado para consolidar vários arquivos de uma mesma pasta:

```m
let
    Fonte = Folder.Files(CaminhoBase),
    FiltraExcel = Table.SelectRows(Fonte, each Text.EndsWith([Extension], ".xlsx")),
    RemoveTemporarios = Table.SelectRows(FiltraExcel, each not Text.StartsWith([Name], "~$"))
in
    RemoveTemporarios
```

> O filtro `~$` remove arquivos temporários que o Excel cria quando a planilha está aberta — sem isso a atualização dá erro.

### Onde alterar depois de publicar
No **Power BI Service**: `Configurações do semântico modelo (dataset)` → **Parâmetros**. Dá para trocar o valor sem reabrir o `.pbix`.

### Outros usos comuns de parâmetros
- **Ambiente (Dev / Homologação / Produção):** usar `Lista de valores` para alternar a fonte de dados
- **Filtro de período:** parâmetro de data para carregar só os últimos X meses e deixar o desenvolvimento mais rápido
- **`RangeStart` e `RangeEnd`:** parâmetros do tipo `Data/Hora` obrigatórios para configurar **atualização incremental**
- **Nome de servidor/banco:** evita reescrever conexão ao migrar de servidor

---

## 02 — Tabela calendário (dim_calendario) com parâmetros de data

### Por que preciso disso
As funções de inteligência temporal do DAX (`TOTALYTD`, `SAMEPERIODLASTYEAR`, `DATEADD`...) só funcionam corretamente com uma **tabela de datas contínua**, sem buracos e com todos os dias do período. A data que já existe na tabela fato não serve: ela tem lacunas (dias sem venda) e vira contexto de filtro errado.

### Passo 1 — Criar os parâmetros de data

Em **Gerenciar Parâmetros** → **Novo**, criar dois parâmetros:

| Nome | Tipo | Valor atual |
|------|------|-------------|
| `MinDate` | Data | `01/01/2020` |
| `MaxDate` | Data | `31/12/2024` |

> Assim eu controlo o intervalo do calendário sem mexer no código. Para mudar o período, altero só o parâmetro.

### Passo 2 — Nova consulta em branco

**Página Inicial** → **Nova Fonte** → **Consulta Nula** → **Editor Avançado**.

### Passo 3 — O código M completo

```m
let
    // 1. Puxa as datas definidas nos parâmetros
    DataInicial = MinDate,
    DataFinal   = MaxDate,

    // 2. Conta quantos dias existem no intervalo
    //    O +1 é necessário para incluir o último dia (senão a data final fica de fora)
    QtdeDias = Duration.Days(DataFinal - DataInicial) + 1,

    // 3. Gera a lista de datas, uma por dia
    //    #duration(dias, horas, minutos, segundos) -> define o passo entre uma data e outra
    ListaDatas = List.Dates(DataInicial, QtdeDias, #duration(1, 0, 0, 0)),

    // 4. Converte a lista em tabela
    TabelaDatas = Table.FromList(ListaDatas, Splitter.SplitByNothing(), {"Data"}, null, ExtraValues.Error),

    // 5. Tipa a coluna como Data
    TipoAjustado = Table.TransformColumnTypes(TabelaDatas, {{"Data", type date}})
in
    TipoAjustado
```

Renomear a consulta para **`dim_calendario`**.

**O que cada peça faz:**

| Comando | Função |
|---------|--------|
| `Duration.Days(fim - inicio)` | Diferença em dias entre as duas datas |
| `+ 1` | Inclui o último dia do período |
| `List.Dates(inicio, qtde, passo)` | Gera a lista de datas |
| `#duration(1,0,0,0)` | Passo de 1 dia (dia, hora, minuto, segundo) |
| `Table.FromList(...)` | Transforma a lista em tabela (botão **Para a Tabela**) |

### Passo 4 — Adicionar as colunas de apoio

Pelo menu **Adicionar Coluna** → **Data**, ou direto no M:

```m
    Ano          = Table.AddColumn(TipoAjustado, "Ano", each Date.Year([Data]), Int64.Type),
    NumMes       = Table.AddColumn(Ano, "NumMes", each Date.Month([Data]), Int64.Type),
    NomeMes      = Table.AddColumn(NumMes, "NomeMes", each Date.ToText([Data], "MMM", "pt-BR"), type text),
    Trimestre    = Table.AddColumn(NomeMes, "Trimestre", each "T" & Text.From(Date.QuarterOfYear([Data])), type text),
    AnoMes       = Table.AddColumn(Trimestre, "AnoMes", each Date.ToText([Data], "yyyy/MM"), type text),
    AnoMesOrdem  = Table.AddColumn(AnoMes, "AnoMesOrdem", each Date.Year([Data]) * 100 + Date.Month([Data]), Int64.Type),
    DiaSemana    = Table.AddColumn(AnoMesOrdem, "DiaSemana", each Date.ToText([Data], "ddd", "pt-BR"), type text),
    NumDiaSemana = Table.AddColumn(DiaSemana, "NumDiaSemana", each Date.DayOfWeek([Data], Day.Monday) + 1, Int64.Type)
```

> **Colunas de ordenação são obrigatórias.** Sem `NumMes` e `AnoMesOrdem`, o Power BI ordena "Abr, Ago, Dez..." em ordem alfabética. No Power BI Desktop: selecionar a coluna `NomeMes` → **Classificar por Coluna** → `NumMes`.

### Passo 5 — Marcar como tabela de datas

No Power BI Desktop: clicar na tabela `dim_calendario` → guia **Design da Tabela** → **Marcar como tabela de datas** → escolher a coluna `Data`.

Sem esse passo, as funções de inteligência temporal podem retornar resultados errados.

### Variação: intervalo dinâmico pela tabela fato

Em vez de fixar `MaxDate` no parâmetro, dá para calcular a partir dos próprios dados — o calendário cresce sozinho a cada atualização:

```m
    DataInicial = Date.StartOfYear(List.Min(fVendas[DataVenda])),
    DataFinal   = Date.EndOfYear(List.Max(fVendas[DataVenda])),
```

> `StartOfYear` / `EndOfYear` garantem que o calendário sempre comece em 01/jan e termine em 31/dez — importante para cálculos de ano completo (YTD).

### Alternativa em DAX

Também dá para criar a tabela por DAX, com **Modelagem** → **Nova Tabela**:

```dax
dim_calendario = CALENDAR(DATE(2020,1,1), DATE(2024,12,31))
```

ou `CALENDARAUTO()`, que varre todas as colunas de data do modelo automaticamente.

**Quando usar cada uma:** prefiro o Power Query (M) porque a tabela vira parte do ETL, é reaproveitável entre projetos e permite dobra de consulta. DAX serve bem para protótipo rápido.

---

## ✅ Boas práticas que vou seguindo

- **Nomear tudo:** cada etapa aplicada com nome descritivo (`RemoveNulos` em vez de `Linhas Filtradas1`)
- **Organizar em grupos:** no Power Query, criar pastas `00 - Parâmetros`, `01 - Fontes`, `02 - Transformações`, `03 - Tabelas Finais`
- **Desabilitar carga** (`Habilitar carga` desmarcado) em consultas intermediárias — elas não precisam virar tabela no modelo
- **Tipar as colunas** logo no início e conferir de novo no fim
- **Remover colunas não usadas** o quanto antes: menos memória, atualização mais rápida
- **Schema em estrela:** tabelas fato no centro, dimensões ao redor — evitar tabelão único
- **Tabela calendário própria**, marcada como Tabela de Datas, para as funções de inteligência temporal funcionarem
- **Medidas em vez de colunas calculadas** sempre que possível (colunas ocupam memória no modelo)

---

## 📌 Próximos tópicos para estudar

- [ ] Coluna de feriados na `dim_calendario`
- [ ] Flags úteis: `EhDiaUtil`, `EhMesAtual`, `EhAnoAtual`
- [ ] Funções personalizadas no Power Query (`(parametro) => let ... in ...`)
- [ ] Mesclar (Merge) vs Anexar (Append)
- [ ] Coluna condicional e Coluna personalizada
- [ ] Dobra de consulta (Query Folding)
- [ ] `CALCULATE`, `FILTER` e contexto de filtro
- [ ] Funções de inteligência temporal (`SAMEPERIODLASTYEAR`, `DATEADD`, `TOTALYTD`)
- [ ] Variáveis em DAX (`VAR` / `RETURN`)
- [ ] `SUMX` e demais funções iteradoras
- [ ] Atualização incremental
- [ ] Row-Level Security (RLS)

---

## 📚 Referências

- [Documentação oficial do Power BI](https://learn.microsoft.com/pt-br/power-bi/)
- [Referência da linguagem M](https://learn.microsoft.com/pt-br/powerquery-m/)
- [Referência da linguagem DAX](https://learn.microsoft.com/pt-br/dax/)
- [DAX Guide](https://dax.guide/)

---

*Repositório em construção — atualizado conforme os estudos avançam.*
