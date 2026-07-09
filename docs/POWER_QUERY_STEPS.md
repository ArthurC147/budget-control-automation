# Power Query — Combinar Actual + Budget e Calcular Variância
**Projeto:** budget-control-automation

Power Query é a ferramenta de transformação de dados dentro do Excel e do
Power BI. Diferente do VBA (que é código), o Power Query funciona através
de uma interface de cliques — mas por trás, ele gera um código chamado
**M language**, que também pode ser editado diretamente.

Este guia mostra as duas formas: pelos cliques (mais fácil para iniciantes)
e o código M equivalente (para quem quiser entender o que acontece por trás).

---

## Por que usar Power Query além da macro VBA?

A macro VBA já calcula a variância dentro do Excel. O Power Query serve
para um propósito complementar: **preparar os dados de forma repetível
antes de qualquer análise** — por exemplo, ao conectar o Power BI
diretamente aos dados, sem depender do Excel estar aberto.

No mercado, é comum ver os dois sendo usados juntos: VBA para automação
de processos dentro do Excel, Power Query para ETL (Extract, Transform, Load)
que alimenta dashboards.

---

## PASSO A PASSO — VIA CLIQUES (Excel)

### 1. Abrir o Power Query
1. Abra `budget_control_data.xlsx`
2. Vá na aba **Dados** (Data) → clique em **Obter Dados** (Get Data) →
   **De Outras Fontes** → **Consulta em Branco** (Blank Query)

### 2. Carregar a aba Actual
1. No editor do Power Query, vá em **Início** → **Nova Fonte** → **Excel**
2. Selecione o próprio arquivo `budget_control_data.xlsx`
3. Marque a tabela/aba **Actual** e clique em **Carregar**

### 3. Carregar a aba Budget
Repita o processo acima, mas selecionando a aba **Budget**.

### 4. Combinar as duas consultas (Merge)
1. Com a consulta **Actual** selecionada, vá em **Início** → **Mesclar Consultas** (Merge Queries)
2. Selecione **Budget** como a segunda tabela
3. Clique nas colunas **Month** e **Category** em ambas as tabelas
   (isso define a chave de cruzamento — equivalente ao PROCV/VLOOKUP)
4. Tipo de junção: **Interna (Inner)** — só mantém combinações que existem nos dois lados
5. Clique em **OK**

### 5. Expandir a coluna combinada
1. Uma nova coluna aparece com o resultado do Budget agrupado
2. Clique no ícone de expandir (duas setas) no cabeçalho dessa coluna
3. Marque apenas **Budget_Sales** e clique em **OK**

### 6. Criar a coluna de variância
1. Vá em **Adicionar Coluna** → **Coluna Personalizada** (Custom Column)
2. Nome da coluna: `Variance_R$`
3. Fórmula: `[Actual_Sales] - [Budget_Sales]`
4. Clique em **OK**

### 7. Criar a coluna de variância percentual
1. **Adicionar Coluna** → **Coluna Personalizada** novamente
2. Nome: `Variance_Pct`
3. Fórmula: `([Actual_Sales] - [Budget_Sales]) / [Budget_Sales] * 100`

### 8. Carregar o resultado
Clique em **Fechar e Carregar** (Close & Load) — os dados combinados
aparecem em uma nova aba da planilha, prontos para o Power BI ou para
qualquer análise adicional.

---

## O CÓDIGO M EQUIVALENTE (gerado automaticamente pelos cliques acima)

Se você abrir o **Editor Avançado** (Advanced Editor) dentro do Power
Query depois de fazer os passos acima, vai ver um código parecido com este:

```m
let
    // Origem: carrega a tabela Actual do arquivo Excel
    Origem = Excel.Workbook(File.Contents("budget_control_data.xlsx"), null, true),
    Actual_Table = Origem{[Item="Actual",Kind="Sheet"]}[Data],

    // Promove a primeira linha como cabeçalho (nomes das colunas)
    ActualComCabecalho = Table.PromoteHeaders(Actual_Table, [PromoteAllScalars=true]),

    // Carrega a tabela Budget da mesma forma
    Budget_Table = Origem{[Item="Budget",Kind="Sheet"]}[Data],
    BudgetComCabecalho = Table.PromoteHeaders(Budget_Table, [PromoteAllScalars=true]),

    // Mescla (merge) as duas tabelas usando Month + Category como chave
    // JoinKind.Inner = mantém só as combinações que existem nos dois lados
    Mesclado = Table.NestedJoin(
        ActualComCabecalho, {"Month", "Category"},
        BudgetComCabecalho, {"Month", "Category"},
        "BudgetExpandido", JoinKind.Inner
    ),

    // Expande a coluna aninhada, trazendo só o Budget_Sales
    Expandido = Table.ExpandTableColumn(
        Mesclado, "BudgetExpandido", {"Budget_Sales"}, {"Budget_Sales"}
    ),

    // Adiciona a coluna de variância em valor absoluto (R$)
    ComVariancia = Table.AddColumn(
        Expandido, "Variance_R$",
        each [Actual_Sales] - [Budget_Sales], type number
    ),

    // Adiciona a coluna de variância percentual
    ComVariânciaPct = Table.AddColumn(
        ComVariancia, "Variance_Pct",
        each ([Actual_Sales] - [Budget_Sales]) / [Budget_Sales] * 100, type number
    )
in
    ComVariânciaPct
```

### Explicação linha por linha do código M

| Trecho | O que faz |
|---|---|
| `let ... in` | Estrutura básica do M — declara passos (`let`) e retorna o resultado final (`in`) |
| `Excel.Workbook(File.Contents(...))` | Abre o arquivo Excel e lê seu conteúdo bruto |
| `Table.PromoteHeaders` | Transforma a primeira linha de dados em nomes de coluna |
| `Table.NestedJoin` | Equivalente a um `JOIN` de SQL ou um PROCV — cruza duas tabelas por uma chave comum |
| `JoinKind.Inner` | Tipo de cruzamento: mantém só o que existe nas duas tabelas (como `INNER JOIN` em SQL) |
| `Table.ExpandTableColumn` | "Abre" a coluna que veio aninhada do merge, trazendo só os campos que interessam |
| `Table.AddColumn` | Cria uma nova coluna calculada — equivalente a uma coluna calculada no Excel ou um `SELECT ... AS` em SQL |
| `each [Actual_Sales] - [Budget_Sales]` | Fórmula aplicada linha por linha (each = "para cada linha") |

Note a semelhança com o SQL do projeto 2: `Table.NestedJoin` com `JoinKind.Inner`
é conceitualmente o mesmo que um `INNER JOIN` em SQL. Saber os dois idiomas
(M e SQL) e reconhecer que resolvem o mesmo problema de formas diferentes
é exatamente o tipo de conhecimento transversal que entrevistas técnicas
de Analytics costumam testar.

---

## Onde usar isso no README do projeto

No README do projeto, uma screenshot do Power Query com as etapas
("Applied Steps") visíveis no painel direito é uma prova visual forte de
que você sabe usar a ferramenta — cole um print dessa tela em
`docs/screenshots/power_query_steps.png`.
