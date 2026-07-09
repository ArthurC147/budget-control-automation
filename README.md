# Budget Control Automation

![Excel](https://img.shields.io/badge/Excel%20VBA-217346?style=flat-square&logo=microsoftexcel&logoColor=white)
![Power Query](https://img.shields.io/badge/Power%20Query-F2C811?style=flat-square&logo=powerbi&logoColor=black)
![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white)
![Status](https://img.shields.io/badge/Status-Active-brightgreen?style=flat-square)

> Automação de controle orçamentário (Budget vs Actual) usando Excel VBA, Power Query e Python — eliminando o fechamento manual mensal de variância financeira.

---

## Contexto de negócio

Todo mês, times financeiros comparam o que foi **gasto/vendido de fato**
(Actual) contra o que **estava planejado** (Budget). Sem automação, isso
significa abrir duas planilhas, cruzar categoria por categoria manualmente,
calcular a variância na mão e formatar tudo antes de apresentar para a liderança.

Este projeto automatiza esse processo com uma macro VBA que cruza os dados,
calcula a variância e já entrega o relatório formatado — a mesma lógica
aplicada no controle orçamentário da **EngePro Jr. Consulting**, que resultou
em **+60% de acurácia financeira** através de automação VBA e padronização de processos.

---

## Problema → Solução → Resultado

| | |
|---|---|
| **Problema** | Comparação manual de Actual vs Budget, categoria por categoria, sujeita a erro humano e sem padronização visual |
| **Solução** | Macro VBA que cruza automaticamente as duas bases, calcula variância % e aplica formatação condicional (verde/amarelo/vermelho) |
| **Resultado** | Relatório de variância gerado em segundos, com desvio médio identificado por categoria e priorização visual imediata |

---

## Dataset

**Fonte:** [Sample Superstore — Kaggle](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final)

Dataset de vendas de varejo usado aqui como base do **Actual** (realizado).
O **Budget** (orçamento) é construído a partir de uma regra de negócio real:
*média móvel dos últimos 3 meses × meta de crescimento por categoria* — a
mesma lógica usada em ferramentas corporativas de FP&A (Financial Planning
& Analysis).

```
data/
├── raw/
│   └── Sample_Superstore.csv       ← original do Kaggle
└── processed/
    ├── budget_control_data.xlsx    ← 3 abas: Actual, Budget, Variance_Report
    └── budget_kpi_summary.csv      ← KPIs finais para o Power BI
```

---

## Como o projeto funciona — 3 camadas de automação

### Camada 1 — Python (preparação dos dados)
`01_coleta.ipynb` e `02_tratamento.ipynb` agregam as vendas por mês/categoria
e constroem o orçamento com a lógica de média móvel, exportando um Excel
de 3 abas pronto para a macro processar.

### Camada 2 — Excel VBA (a automação central)
`RunBudgetControl.bas` é uma macro que:
1. Lê as abas **Actual** e **Budget**
2. Cruza os dois conjuntos por Mês + Categoria (equivalente a um PROCV/VLOOKUP, mas em código)
3. Calcula a variância em R$ e em %
4. Aplica formatação condicional automática (verde ≤10%, amarelo ≤25%, vermelho >25%)
5. Escreve tudo na aba **Variance_Report**, já formatado

### Camada 3 — Power Query (ETL alternativo/complementar)
`POWER_QUERY_STEPS.md` documenta como fazer o mesmo cruzamento via Power
Query (M language) — útil para conectar diretamente ao Power BI sem
depender da macro rodar antes.

### Camada 4 — Power BI (visualização)
O arquivo `churn_kpi_summary.csv` alimenta um dashboard com a variância
por categoria e por mês.

---

## Estrutura do projeto

```
budget-control-automation/
├── data/
│   ├── raw/
│   │   └── Sample_Superstore.csv
│   └── processed/
│       ├── budget_control_data.xlsx
│       └── budget_kpi_summary.csv
├── notebooks/
│   ├── 01_coleta.ipynb
│   ├── 02_tratamento.ipynb
│   └── 03_analise.ipynb
├── vba/
│   └── RunBudgetControl.bas
├── docs/
│   ├── POWER_QUERY_STEPS.md
│   └── screenshots/
│       ├── 01_sales_by_category.png
│       ├── budget_variance_trend.png
│       ├── budget_status_distribution.png
│       └── budget_avg_deviation.png
├── dashboard/
│   └── budget_dashboard.pbix
├── requirements.txt
└── README.md
```

---

## Como executar

### 1. Preparar os dados (Python)
```bash
git clone https://github.com/ArthurC147/budget-control-automation.git
cd budget-control-automation
pip install -r requirements.txt
```
Baixe `Sample_Superstore.csv` do Kaggle e coloque em `data/raw/`.

Rode os notebooks em ordem:
```bash
jupyter notebook
# 01_coleta.ipynb → 02_tratamento.ipynb
```
Isso gera `data/processed/budget_control_data.xlsx`.

### 2. Rodar a automação VBA
1. Abra `budget_control_data.xlsx` no Excel
2. Pressione **Alt+F11** → **Inserir → Módulo**
3. Cole o conteúdo de `vba/RunBudgetControl.bas`
4. Feche o editor (**Alt+Q**) → pressione **Alt+F8** → selecione `RunBudgetControl` → **Executar**
5. A aba **Variance_Report** será preenchida automaticamente

### 3. Gerar os gráficos finais (Python)
Com o Excel salvo após rodar a macro, execute:
```bash
jupyter notebook
# 03_analise.ipynb
```

### 4. Power Query (opcional)
Siga `docs/POWER_QUERY_STEPS.md` para replicar o cruzamento via Power Query.

---

## Dependências

```
pandas>=2.0.0
numpy>=1.24.0
matplotlib>=3.7.0
seaborn>=0.12.0
openpyxl>=3.1.0
jupyter>=1.0.0
```

---

## O que eu aprendi

- VBA não tem os recursos de alto nível do Pandas (`.groupby()`, `.merge()`) —
  cruzar duas tabelas exige escrever o loop de busca manualmente, o que
  aprofundou meu entendimento de como um PROCV/VLOOKUP funciona por trás
- Orçamento como "média móvel × meta de crescimento" é mais defensável
  que um número fixo — fica claro de onde cada valor vem, o que facilita
  auditoria e explicação para stakeholders
- Power Query e SQL resolvem o mesmo problema (cruzar tabelas) com sintaxes
  diferentes — `Table.NestedJoin` com `JoinKind.Inner` é conceitualmente
  idêntico a um `INNER JOIN` em SQL
- Formatação condicional automática (verde/amarelo/vermelho) transforma
  uma tabela de números em uma ferramenta de priorização visual imediata

---

## Autor

**Arthur Cardoso** — Industrial Engineering @ UFPR · Business & Customer Success Intern @ Telefônica Vivo

[![LinkedIn](https://img.shields.io/badge/LinkedIn-0A66C2?style=flat-square&logo=linkedin&logoColor=white)](https://linkedin.com/in/arthur-cardoso-b3b1ba1ab)
[![GitHub](https://img.shields.io/badge/GitHub-181717?style=flat-square&logo=github&logoColor=white)](https://github.com/ArthurC147)
