Attribute VB_Name = "Module1"
' ============================================================================
' RunBudgetControl.bas
' Projeto: budget-control-automation
'
' O QUE ESTA MACRO FAZ:
'   1. Lê os dados das abas "Actual" e "Budget"
'   2. Cruza os dois (mesma lógica de um PROCV/VLOOKUP, mas em código)
'   3. Calcula a variância (Actual - Budget) e a variância percentual
'   4. Aplica formatação condicional (verde = dentro da meta, vermelho = estourou)
'   5. Escreve tudo automaticamente na aba "Variance_Report"
'
' COMO INSTALAR:
'   1. Abra o arquivo budget_control_data.xlsx no Excel
'   2. Pressione Alt+F11 para abrir o Editor VBA
'   3. No menu: Inserir → Módulo (Insert → Module)
'   4. Cole todo o código abaixo dentro do módulo
'   5. Feche o Editor VBA (Alt+Q)
'   6. Na planilha, pressione Alt+F8, selecione "RunBudgetControl" e clique em Executar
' ============================================================================

Sub RunBudgetControl()

    ' ── Declaração de variáveis ──────────────────────────────────────────
    ' Em VBA, toda variável precisa ser declarada com um tipo antes de usar.
    ' Dim = "Dimension" = reserva espaço de memória para a variável.

    Dim wsActual As Worksheet      ' referência para a aba "Actual"
    Dim wsBudget As Worksheet      ' referência para a aba "Budget"
    Dim wsReport As Worksheet      ' referência para a aba "Variance_Report"

    Dim lastRowActual As Long      ' última linha com dado na aba Actual
    Dim lastRowBudget As Long      ' última linha com dado na aba Budget
    Dim i As Long                  ' contador do loop principal
    Dim j As Long                  ' contador do loop de busca (cruzamento)

    Dim monthActual As String      ' mês da linha atual em Actual
    Dim catActual As String        ' categoria da linha atual em Actual
    Dim salesActual As Double      ' valor de vendas realizado
    Dim profitActual As Double     ' lucro realizado

    Dim monthBudget As String
    Dim catBudget As String
    Dim budgetValue As Double

    Dim variance As Double         ' Actual - Budget em valor absoluto (R$)
    Dim variancePct As Double      ' variação percentual
    Dim found As Boolean           ' controla se encontrou o par Month+Category no Budget

    Dim reportRow As Long          ' linha atual de escrita na aba Variance_Report

    ' ── Conectar às abas (evita erro se o nome estiver diferente) ────────
    On Error Resume Next
    Set wsActual = ThisWorkbook.Sheets("Actual")
    Set wsBudget = ThisWorkbook.Sheets("Budget")
    Set wsReport = ThisWorkbook.Sheets("Variance_Report")
    On Error GoTo 0

    If wsActual Is Nothing Or wsBudget Is Nothing Or wsReport Is Nothing Then
        MsgBox "Erro: confirme que as abas 'Actual', 'Budget' e 'Variance_Report' existem.", vbCritical
        Exit Sub
    End If

    ' ── Limpar a aba de relatório antes de gerar de novo ──────────────────
    ' Isso garante que rodar a macro várias vezes não duplica os dados
    wsReport.Cells.Clear

    ' ── Escrever o cabeçalho da aba Variance_Report ───────────────────────
    wsReport.Range("A1").Value = "Month"
    wsReport.Range("B1").Value = "Category"
    wsReport.Range("C1").Value = "Actual_Sales"
    wsReport.Range("D1").Value = "Budget_Sales"
    wsReport.Range("E1").Value = "Variance_R$"
    wsReport.Range("F1").Value = "Variance_%"
    wsReport.Range("G1").Value = "Status"

    ' Formatar o cabeçalho: fundo azul, texto branco, negrito
    With wsReport.Range("A1:G1")
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' ── Descobrir até qual linha cada aba tem dados ───────────────────────
    ' .Cells(Rows.Count, 1).End(xlUp).Row = "vá até a última linha da planilha
    ' na coluna A, e suba (End xlUp) até encontrar a última célula preenchida"
    ' Essa é a forma padrão em VBA de descobrir dinamicamente o tamanho dos dados,
    ' sem precisar digitar um número fixo de linhas.

    lastRowActual = wsActual.Cells(wsActual.Rows.Count, 1).End(xlUp).Row
    lastRowBudget = wsBudget.Cells(wsBudget.Rows.Count, 1).End(xlUp).Row

    reportRow = 2   ' começamos a escrever os dados na linha 2 (linha 1 é o cabeçalho)

    ' ── Loop principal: percorre cada linha da aba Actual ─────────────────
    ' For i = 2 To lastRowActual = "repita para cada linha, começando na 2
    ' (pulando o cabeçalho) até a última linha com dado"

    For i = 2 To lastRowActual

        monthActual = wsActual.Cells(i, 1).Value    ' coluna A = Month
        catActual = wsActual.Cells(i, 2).Value       ' coluna B = Category
        salesActual = wsActual.Cells(i, 3).Value     ' coluna C = Actual_Sales
        profitActual = wsActual.Cells(i, 4).Value    ' coluna D = Actual_Profit

        found = False

        ' ── Loop de busca: procura o Mês+Categoria correspondente no Budget ──
        ' Isso é equivalente a um PROCV (VLOOKUP), mas feito manualmente em VBA,
        ' o que dá mais controle (ex: podemos comparar 2 colunas ao mesmo tempo:
        ' Month E Category, algo que um VLOOKUP simples não faz sozinho)

        For j = 2 To lastRowBudget
            monthBudget = wsBudget.Cells(j, 1).Value
            catBudget = wsBudget.Cells(j, 2).Value

            If monthBudget = monthActual And catBudget = catActual Then
                budgetValue = wsBudget.Cells(j, 3).Value
                found = True
                Exit For   ' encontrou o par certo, não precisa continuar procurando
            End If
        Next j

        ' ── Se encontrou o orçamento correspondente, calcula a variância ────
        If found Then
            variance = salesActual - budgetValue

            ' Evita erro de divisão por zero se o orçamento for 0
            If budgetValue <> 0 Then
                variancePct = (variance / budgetValue) * 100
            Else
                variancePct = 0
            End If

            ' ── Escrever a linha no relatório ──────────────────────────────
            wsReport.Cells(reportRow, 1).Value = monthActual
            wsReport.Cells(reportRow, 2).Value = catActual
            wsReport.Cells(reportRow, 3).Value = salesActual
            wsReport.Cells(reportRow, 4).Value = budgetValue
            wsReport.Cells(reportRow, 5).Value = Round(variance, 2)
            wsReport.Cells(reportRow, 6).Value = Round(variancePct, 1)

            ' ── Regra de negócio: classificar o status ─────────────────────
            If variancePct >= 0 Then
                wsReport.Cells(reportRow, 7).Value = "Acima do orçamento"
            Else
                wsReport.Cells(reportRow, 7).Value = "Abaixo do orçamento"
            End If

            ' ── Formatação condicional: colore a célula de variância ───────
            ' Verde: variação entre -10% e +15% (dentro do aceitável)
            ' Amarelo: variação moderada (atenção)
            ' Vermelho: desvio grande (>25% para qualquer lado) — precisa de ação

            If Abs(variancePct) <= 10 Then
                wsReport.Cells(reportRow, 6).Interior.Color = RGB(198, 239, 206)   ' verde claro
            ElseIf Abs(variancePct) <= 25 Then
                wsReport.Cells(reportRow, 6).Interior.Color = RGB(255, 235, 156)   ' amarelo claro
            Else
                wsReport.Cells(reportRow, 6).Interior.Color = RGB(255, 199, 206)   ' vermelho claro
            End If

            reportRow = reportRow + 1   ' avança para a próxima linha do relatório
        End If

    Next i

    ' ── Ajustar largura das colunas automaticamente ───────────────────────
    wsReport.Columns("A:G").AutoFit

    ' ── Adicionar um resumo executivo no topo, algumas linhas abaixo dos dados ──
    Dim summaryRow As Long
    summaryRow = reportRow + 2

    wsReport.Cells(summaryRow, 1).Value = "RESUMO EXECUTIVO"
    wsReport.Cells(summaryRow, 1).Font.Bold = True

    wsReport.Cells(summaryRow + 1, 1).Value = "Total de linhas analisadas:"
    wsReport.Cells(summaryRow + 1, 2).Value = reportRow - 2

    wsReport.Cells(summaryRow + 2, 1).Value = "Desvio médio absoluto (%):"
    ' NOTA TÉCNICA: AVERAGE(ABS(...)) exigiria ser digitada como "fórmula
    ' matricial" (Ctrl+Shift+Enter) em versões do Excel sem array dinâmico —
    ' sem isso, o resultado sairia errado (só consideraria a primeira célula).
    ' SUMPRODUCT resolve isso nativamente, sem precisar de entrada especial:
    wsReport.Cells(summaryRow + 2, 2).Value = _
        "=ROUND(SUMPRODUCT(ABS(F2:F" & (reportRow - 1) & "))/COUNT(F2:F" & (reportRow - 1) & "),1)"

    ' ── Avisar o usuário que terminou ──────────────────────────────────────
    MsgBox "Relatório de variância gerado com sucesso!" & vbNewLine & _
           reportRow - 2 & " linhas processadas.", vbInformation, "Budget Control Automation"

End Sub
