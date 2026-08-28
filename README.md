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
