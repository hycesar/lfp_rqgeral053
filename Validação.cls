VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Validação"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Calcula_click()
    Dim VarVar1 As Double, VarVar2 As Double, FCalc As Double, FCrit As Double, Aviso As Boolean
    Aviso = False

    'Resume Next
    On Error GoTo Erro
    
    If AddIns("Ferramentas de Análise - VBA").Installed = False And AddIns("Ferramentas de Análise").Installed = False Then
       A = MsgBox("Não foi possível carregar o suplemento 'Análise de Dados - VBA'. Faça a instalação manualmente" & vbNewLine & vbNewLine & "                                       Em caso de Dúvida contacte COBPF", vbCritical, "Erro Crítico - Impossível Continuar!")
       Else
        'Prólogo
            'UnProtect
            ActiveSheet.Unprotect Password:="0"
            'Instalar Suplementos
            AddIns("Ferramentas de Análise - VBA").Installed = True
            'Speed Up
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False
            Application.DisplayStatusBar = False
            Application.EnableEvents = False
            ActiveSheet.DisplayPageBreaks = False

        'Linearidade
            If Range("A10").FormulaR1C1 = 5 Then
            'Gráfico
                ActiveSheet.ChartObjects("gfLinearidadeM").Activate
                ActiveChart.SeriesCollection(1).XValues = "='Validação'!$B$12:$B$16"
                ActiveChart.SeriesCollection(1).Values = "='Validação'!$I$12:$I$16"
                ActiveSheet.ChartObjects("gfLinearidade3").Activate
                ActiveChart.SeriesCollection(1).XValues = "='Validação'!$B$12:$B$16"
                ActiveChart.SeriesCollection(1).Values = "='Validação'!$H$12:$H$16"
                ActiveSheet.ChartObjects("gfLinearidade2").Activate
                ActiveChart.SeriesCollection(1).XValues = "='Validação'!$B$12:$B$16"
                ActiveChart.SeriesCollection(1).Values = "='Validação'!$G$12:$G$16"
                ActiveSheet.ChartObjects("gfLinearidade1").Activate
                ActiveChart.SeriesCollection(1).XValues = "='Validação'!$B$12:$B$16"
                ActiveChart.SeriesCollection(1).Values = "='Validação'!$F$12:$F$16"
            'Estatistica
                Range("M13:N13").FormulaArray = "=LINEST(R[-1]C[-7]:R[3]C[-7],R[-1]C[-11]:R[3]C[-11],TRUE,FALSE)"
                Range("M14:N14").FormulaArray = "=LINEST(R[-2]C[-6]:R[2]C[-6],R[-2]C[-11]:R[2]C[-11],TRUE,FALSE)"
                Range("M15:N15").FormulaArray = "=LINEST(R[-3]C[-5]:R[1]C[-5],R[-3]C[-11]:R[1]C[-11],TRUE,FALSE)"
                Range("M17:N17").FormulaArray = "=LINEST(R[-5]C[-4]:R[-1]C[-4],R[-5]C[-11]:R[-1]C[-11],TRUE,FALSE)"
                Range("O13").FormulaR1C1 = "=IF(R[3]C[-9]>0,SQRT(RSQ(R[-1]C[-9]:R[3]C[-9],R[-1]C[-13]:R[3]C[-13])),0)"
                Range("O14").FormulaR1C1 = "=IF(R[2]C[-8]>0,SQRT(RSQ(R[-2]C[-8]:R[2]C[-8],R[-2]C[-13]:R[2]C[-13])),0)"
                Range("O15").FormulaR1C1 = "=IF(R[1]C[-7]>0,SQRT(RSQ(R[-3]C[-7]:R[1]C[-7],R[-3]C[-13]:R[1]C[-13])),0)"
                Range("O17").FormulaR1C1 = "=IF(R[-1]C[-6]>0,SQRT(RSQ(R[-5]C[-6]:R[-1]C[-6],R[-5]C[-13]:R[-1]C[-13])),0)"
            'Anova
                Range("A19:G33").Clear
                Application.Run "ATPVBAEN.XLAM!Anova1", ActiveSheet.Range("$F$12:$H$16"), ActiveSheet.Range("$A$19"), "C", False, 0.05
                Range("G29:G30").Cut Destination:=Range("E32")
                With Range("E33").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                Range("G33").Clear
                Range("A29").FormulaR1C1 = "Variação"
                Range("A30").FormulaR1C1 = "Intragrupos"
                Range("A31").FormulaR1C1 = "Intergrupos"
                Range("A21:E25").Cut Destination:=Range("A20")
                Range("A28:F33").Cut Destination:=Range("A25")
            ElseIf Range("A10").FormulaR1C1 = 6 Then
            'Gráfico
                ActiveSheet.ChartObjects("gfLinearidadeM").Activate
                ActiveChart.SeriesCollection(1).XValues = "='Validação'!$B$12:$B$17"
                ActiveChart.SeriesCollection(1).Values = "='Validação'!$I$12:$I$17"
                ActiveSheet.ChartObjects("gfLinearidade3").Activate
                ActiveChart.SeriesCollection(1).XValues = "='Validação'!$B$12:$B$17"
                ActiveChart.SeriesCollection(1).Values = "='Validação'!$H$12:$H$17"
                ActiveSheet.ChartObjects("gfLinearidade2").Activate
                ActiveChart.SeriesCollection(1).XValues = "='Validação'!$B$12:$B$17"
                ActiveChart.SeriesCollection(1).Values = "='Validação'!$G$12:$G$17"
                ActiveSheet.ChartObjects("gfLinearidade1").Activate
                ActiveChart.SeriesCollection(1).XValues = "='Validação'!$B$12:$B$17"
                ActiveChart.SeriesCollection(1).Values = "='Validação'!$F$12:$F$17"
            'Estatistica
                Range("M13:N13").FormulaArray = "=LINEST(R[-1]C[-7]:R[4]C[-7],R[-1]C[-11]:R[4]C[-11],TRUE,FALSE)"
                Range("M14:N14").FormulaArray = "=LINEST(R[-2]C[-6]:R[3]C[-6],R[-2]C[-11]:R[3]C[-11],TRUE,FALSE)"
                Range("M15:N15").FormulaArray = "=LINEST(R[-3]C[-5]:R[2]C[-5],R[-3]C[-11]:R[2]C[-11],TRUE,FALSE)"
                Range("M17:N17").FormulaArray = "=LINEST(R[-5]C[-4]:RC[-4],R[-5]C[-11]:RC[-11],TRUE,FALSE)"
                Range("O13").FormulaR1C1 = "=IF(R[4]C[-9]>0,SQRT(RSQ(R[-1]C[-9]:R[4]C[-9],R[-1]C[-13]:R[4]C[-13])),0)"
                Range("O14").FormulaR1C1 = "=IF(R[3]C[-8]>0,SQRT(RSQ(R[-2]C[-8]:R[3]C[-8],R[-2]C[-13]:R[3]C[-13])),0)"
                Range("O15").FormulaR1C1 = "=IF(R[2]C[-7]>0,SQRT(RSQ(R[-3]C[-7]:R[2]C[-7],R[-3]C[-13]:R[2]C[-13])),0)"
                Range("O17").FormulaR1C1 = "=IF(RC[-6]>0,SQRT(RSQ(R[-5]C[-6]:RC[-6],R[-5]C[-13]:RC[-13])),0)"
            'Anova
                Range("A19:G33").Clear
                Application.Run "ATPVBAEN.XLAM!Anova1", ActiveSheet.Range("$F$12:$H$17"), ActiveSheet.Range("$A$19"), "C", False, 0.05
                Range("G29:G30").Cut Destination:=Range("E32")
                With Range("E33").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                Range("G33").Clear
                Range("A29").FormulaR1C1 = "Variação"
                Range("A30").FormulaR1C1 = "Intragrupos"
                Range("A31").FormulaR1C1 = "Intergrupos"
                Range("A21:E25").Cut Destination:=Range("A20")
                Range("A28:F33").Cut Destination:=Range("A25")
            ElseIf Range("A10").FormulaR1C1 = 7 Then
            'Gráfico
                ActiveSheet.ChartObjects("gfLinearidadeM").Activate
                ActiveChart.SeriesCollection(1).XValues = "='Validação'!$B$12:$B$18"
                ActiveChart.SeriesCollection(1).Values = "='Validação'!$I$12:$I$18"
                ActiveSheet.ChartObjects("gfLinearidade3").Activate
                ActiveChart.SeriesCollection(1).XValues = "='Validação'!$B$12:$B$18"
                ActiveChart.SeriesCollection(1).Values = "='Validação'!$H$12:$H$18"
                ActiveSheet.ChartObjects("gfLinearidade2").Activate
                ActiveChart.SeriesCollection(1).XValues = "='Validação'!$B$12:$B$18"
                ActiveChart.SeriesCollection(1).Values = "='Validação'!$G$12:$G$18"
                ActiveSheet.ChartObjects("gfLinearidade1").Activate
                ActiveChart.SeriesCollection(1).XValues = "='Validação'!$B$12:$B$18"
                ActiveChart.SeriesCollection(1).Values = "='Validação'!$F$12:$F$18"
            'Estatistica
                Range("M13:N13").FormulaArray = "=LINEST(R[-1]C[-7]:R[5]C[-7],R[-1]C[-11]:R[5]C[-11],TRUE,FALSE)"
                Range("M14:N14").FormulaArray = "=LINEST(R[-2]C[-6]:R[4]C[-6],R[-2]C[-11]:R[4]C[-11],TRUE,FALSE)"
                Range("M15:N15").FormulaArray = "=LINEST(R[-3]C[-5]:R[3]C[-5],R[-3]C[-11]:R[3]C[-11],TRUE,FALSE)"
                Range("M17:N17").FormulaArray = "=LINEST(R[-5]C[-4]:R[1]C[-4],R[-5]C[-11]:R[1]C[-11],TRUE,FALSE)"
                Range("O13").FormulaR1C1 = "=IF(R[5]C[-9]>0,SQRT(RSQ(R[-1]C[-9]:R[5]C[-9],R[-1]C[-13]:R[5]C[-13])),0)"
                Range("O14").FormulaR1C1 = "=IF(R[4]C[-8]>0,SQRT(RSQ(R[-2]C[-8]:R[4]C[-8],R[-2]C[-13]:R[4]C[-13])),0)"
                Range("O15").FormulaR1C1 = "=IF(R[3]C[-7]>0,SQRT(RSQ(R[-3]C[-7]:R[3]C[-7],R[-3]C[-13]:R[3]C[-13])),0)"
                Range("O17").FormulaR1C1 = "=IF(R[1]C[-6]>0,SQRT(RSQ(R[-5]C[-6]:R[1]C[-6],R[-5]C[-13]:R[1]C[-13])),0)"
            'Anova
                Range("A19:G33").Clear
                Application.Run "ATPVBAEN.XLAM!Anova1", ActiveSheet.Range("$F$12:$H$18"), ActiveSheet.Range("$A$19"), "C", False, 0.05
                Range("G29:G30").Cut Destination:=Range("E32")
                With Range("E33").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                Range("G33").Clear
                Range("A29").FormulaR1C1 = "Variação"
                Range("A30").FormulaR1C1 = "Intragrupos"
                Range("A31").FormulaR1C1 = "Intergrupos"
                Range("A21:E25").Cut Destination:=Range("A20")
                Range("A28:F33").Cut Destination:=Range("A25")
            Else
                MsgBox "Para Linearidade, selecione número de pontos!"
            End If
        'Linearidade
            'Formato dos dados e suas casas decimais
            If Range("B5").FormulaR1C1 = "Espectro UV" Then
                With Range("C12:E18,I12:I18,M13:N18,E45:F45,C47:C49,G47:G49,C52:C57,E52:E57,A66:B66,C68:C70,F65:F70,C78:C80,A76:B76,F78:H80")
                    .NumberFormat = "0.0000"
                    .Font.Size = 5
                End With
                With Range("C85:C87,E85:E87,G85:G87,M85:N87,C89:C91,E89:E91,G89:G91,C103:C105,E103:E105,G103:G105,M103:N105,C107:C109,E107:E109,G107:G109,C126:C128,E126:E128,G126:G128,M126:N128,C130:C132,E130:E132,G130:G132")
                    .NumberFormat = "0.0000"
                    .Font.Size = 5
                End With
            Else
                With Range("C12:E18,I12:I18,M13:N18,E45:F45,C47:C49,G47:G49,C52:C57,E52:E57,A66:B66,C68:C70,F65:F70,C78:C80,A76:B76,F78:H80")
                    .NumberFormat = "0.0"
                    .Font.Size = 6
                End With
                With Range("C85:C87,E85:E87,G85:G87,M85:N87,C89:C91,E89:E91,G89:G91,C103:C105,E103:E105,G103:G105,M103:N105,C107:C109,E107:E109,G107:G109,C126:C128,E126:E128,G126:G128,M126:N128,C130:C132,E130:E132,G130:G132")
                    .NumberFormat = "0.0"
                    .Font.Size = 6
                End With
            End If
            'Reescrever
            Range("M8:O8").FormulaR1C1 = "=IF(R[19]C[-11]<>0, IF(AND(MIN(R[5]C[2]:R[9]C[2])>0.99, MAX(R[4]C[-2]:R[10]C[-2])<0.05,IF(R[19]C[-8]<=R[22]C[-8],TRUE,FALSE)),""PASSOU"",""FALHA""),""N/A"")"

        'Precisão Intermediária
        If Range("A45").FormulaR1C1 = 3 And WorksheetFunction.CountBlank(Range("G52:H54")) = 0 Then
            'Teste F
            Range("I45:O58").Clear
            Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$G$52:$G$54"), ActiveSheet.Range("$H$52:$H$54"), ActiveSheet.Range("$I$45"), False, 0.05
            'Teste F usado corretamente? Variancia da primeira Variável deve ser maior do que a da segunda
            If Application.International(xlDecimalSeparator) = "." Then
                VarVar1 = Range("J49").FormulaR1C1
                VarVar2 = Range("K49").FormulaR1C1
            Else
                VarVar1 = Replace(Range("J49").FormulaR1C1, ".", ",")
                VarVar2 = Replace(Range("K49").FormulaR1C1, ".", ",")
            End If
            If VarVar2 > VarVar1 Then
                'Limpar
                Range("I45:K54").Clear
                'Refazer o Teste F
                Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$H$52:$H$54"), ActiveSheet.Range("$G$52:$G$54"), ActiveSheet.Range("$I$45"), False, 0.05
            End If
            'Rescrever
            Range("I45").FormulaR1C1 = "Teste-F:2 amostras p/ variância"
            'Definir Tipo de Teste t
            If Application.International(xlDecimalSeparator) = "." Then
                FCalc = Range("J52").FormulaR1C1
                FCrit = Range("J54").FormulaR1C1
            Else
                FCalc = Replace(Range("J52").FormulaR1C1, ".", ",")
                FCrit = Replace(Range("J54").FormulaR1C1, ".", ",")
            End If
            If FCalc <= FCrit Then 'Mesma Variancia
                If VarVar1 >= VarVar2 Then 'Ordem Direta
                    Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$G$52:$G$54"), ActiveSheet.Range("$H$52:$H$54"), ActiveSheet.Range("$L$45"), False, 0.05
                Else 'Ordem Inversa
                    Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$H$52:$H$54"), ActiveSheet.Range("$G$52:$G$54"), ActiveSheet.Range("$L$45"), False, 0.05
                End If
                ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância equivalente"
                'Reescreve Formula
                Range("M42:O42").FormulaR1C1 = "=IF(R[12]C<>0,IF(AND(ABS(R[12]C)<=R[16]C),""PASSOU"",""FALHA""),""N/A"")"
            Else 'Diferente Variancia
                If VarVar1 >= VarVar2 Then 'Ordem Direta
                    Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$G$52:$G$54"), ActiveSheet.Range("$H$52:$H$54"), ActiveSheet.Range("$L$45"), False, 0.05
                Else 'Ordem Inversa
                    Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$H$52:$H$54"), ActiveSheet.Range("$G$52:$G$54"), ActiveSheet.Range("$L$45"), False, 0.05
                End If
                ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância diferente"
                'Reescreve Formula
                Range("M42:O42").FormulaR1C1 = "=IF(R[11]C<>0,IF(AND(ABS(R[11]C)<=R[15]C),""PASSOU"",""FALHA""),""N/A"")"
            End If
        ElseIf Range("A45").FormulaR1C1 = 4 And WorksheetFunction.CountBlank(Range("G52:H55")) = 0 Then
            'Precisão Intermediária
            'Teste F
            Range("I45:O58").Clear
            Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$G$52:$G$55"), ActiveSheet.Range("$H$52:$H$55"), ActiveSheet.Range("$I$45"), False, 0.05
            'Teste F usado corretamente? Variancia da primeira Variável deve ser maior do que a da segunda
            If Application.International(xlDecimalSeparator) = "." Then
                VarVar1 = Range("J49").FormulaR1C1
                VarVar2 = Range("K49").FormulaR1C1
            Else
                VarVar1 = Replace(Range("J49").FormulaR1C1, ".", ",")
                VarVar2 = Replace(Range("K49").FormulaR1C1, ".", ",")
            End If
            If VarVar2 > VarVar1 Then
                'Limpar
                Range("I45:K54").Clear
                'Refazer o Teste F
                Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$H$52:$H$55"), ActiveSheet.Range("$G$52:$G$55"), ActiveSheet.Range("$I$45"), False, 0.05
            End If
            'Rescrever
            Range("I45").FormulaR1C1 = "Teste-F:2 amostras p/ variância"
            'Definir Tipo de Teste t
            If Application.International(xlDecimalSeparator) = "." Then
                FCalc = Range("J52").FormulaR1C1
                FCrit = Range("J54").FormulaR1C1
            Else
                FCalc = Replace(Range("J52").FormulaR1C1, ".", ",")
                FCrit = Replace(Range("J54").FormulaR1C1, ".", ",")
            End If
            If FCalc <= FCrit Then 'Mesma Variancia
                If VarVar1 >= VarVar2 Then 'Ordem Direta
                    Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$G$52:$G$55"), ActiveSheet.Range("$H$52:$H$55"), ActiveSheet.Range("$L$45"), False, 0.05
                Else 'Ordem Inversa
                    Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$H$52:$H$55"), ActiveSheet.Range("$G$52:$G$55"), ActiveSheet.Range("$L$45"), False, 0.05
                End If
                ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância equivalente"
                'Reescreve Formula
                Range("M42:O42").FormulaR1C1 = "=IF(R[12]C<>0,IF(AND(ABS(R[12]C)<=R[16]C),""PASSOU"",""FALHA""),""N/A"")"
            Else 'Diferente Variancia
                If VarVar1 >= VarVar2 Then 'Ordem Direta
                    Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$G$52:$G$55"), ActiveSheet.Range("$H$52:$H$55"), ActiveSheet.Range("$L$45"), False, 0.05
                Else 'Ordem Inversa
                    Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$H$52:$H$55"), ActiveSheet.Range("$G$52:$G$55"), ActiveSheet.Range("$L$45"), False, 0.05
                End If
                ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância diferente"
                'Reescreve Formula
                Range("M42:O42").FormulaR1C1 = "=IF(R[11]C<>0,IF(AND(ABS(R[11]C)<=R[15]C),""PASSOU"",""FALHA""),""N/A"")"
            End If
        ElseIf Range("A45").FormulaR1C1 = 5 And WorksheetFunction.CountBlank(Range("G52:H56")) = 0 Then
            'Precisão Intermediária
            'Teste F
            Range("I45:O58").Clear
            Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$G$52:$G$56"), ActiveSheet.Range("$H$52:$H$56"), ActiveSheet.Range("$I$45"), False, 0.05
            'Teste F usado corretamente? Variancia da primeira Variável deve ser maior do que a da segunda
            If Application.International(xlDecimalSeparator) = "." Then
                VarVar1 = Range("J49").FormulaR1C1
                VarVar2 = Range("K49").FormulaR1C1
            Else
                VarVar1 = Replace(Range("J49").FormulaR1C1, ".", ",")
                VarVar2 = Replace(Range("K49").FormulaR1C1, ".", ",")
            End If
            If VarVar2 > VarVar1 Then
                'Limpar
                Range("I45:K54").Clear
                'Refazer o Teste F
                Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$H$52:$H$56"), ActiveSheet.Range("$G$52:$G$56"), ActiveSheet.Range("$I$45"), False, 0.05
            End If
            'Rescrever
            Range("I45").FormulaR1C1 = "Teste-F:2 amostras p/ variância"
            'Definir Tipo de Teste t
            If Application.International(xlDecimalSeparator) = "." Then
                FCalc = Range("J52").FormulaR1C1
                FCrit = Range("J54").FormulaR1C1
            Else
                FCalc = Replace(Range("J52").FormulaR1C1, ".", ",")
                FCrit = Replace(Range("J54").FormulaR1C1, ".", ",")
            End If
            If FCalc <= FCrit Then 'Mesma Variancia
                If VarVar1 >= VarVar2 Then 'Ordem Direta
                    Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$G$52:$G$56"), ActiveSheet.Range("$H$52:$H$56"), ActiveSheet.Range("$L$45"), False, 0.05
                Else 'Ordem Inversa
                    Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$H$52:$H$56"), ActiveSheet.Range("$G$52:$G$56"), ActiveSheet.Range("$L$45"), False, 0.05
                End If
                ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância equivalente"
                'Reescreve Formula
                Range("M42:O42").FormulaR1C1 = "=IF(R[12]C<>0,IF(AND(ABS(R[12]C)<=R[16]C),""PASSOU"",""FALHA""),""N/A"")"
            Else 'Diferente Variancia
                If VarVar1 >= VarVar2 Then 'Ordem Direta
                    Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$G$52:$G$56"), ActiveSheet.Range("$H$52:$H$56"), ActiveSheet.Range("$L$45"), False, 0.05
                Else 'Ordem Inversa
                    Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$H$52:$H$56"), ActiveSheet.Range("$G$52:$G$56"), ActiveSheet.Range("$L$45"), False, 0.05
                End If
                ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância diferente"
                'Reescreve Formula
                Range("M42:O42").FormulaR1C1 = "=IF(R[11]C<>0,IF(AND(ABS(R[11]C)<=R[15]C),""PASSOU"",""FALHA""),""N/A"")"
            End If
        ElseIf Range("A45").FormulaR1C1 = 6 And WorksheetFunction.CountBlank(Range("G52:H57")) = 0 Then
            'Precisão Intermediária
            'Teste F
            Range("I45:O58").Clear
            Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$G$52:$G$57"), ActiveSheet.Range("$H$52:$H$57"), ActiveSheet.Range("$I$45"), False, 0.05
            'Teste F usado corretamente? Variancia da primeira Variável deve ser maior do que a da segunda
            If Application.International(xlDecimalSeparator) = "." Then
                VarVar1 = Range("J49").FormulaR1C1
                VarVar2 = Range("K49").FormulaR1C1
            Else
                VarVar1 = Replace(Range("J49").FormulaR1C1, ".", ",")
                VarVar2 = Replace(Range("K49").FormulaR1C1, ".", ",")
            End If
            If VarVar2 > VarVar1 Then
                'Limpar
                Range("I45:K54").Clear
                'Refazer o Teste F
                Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$H$52:$H$57"), ActiveSheet.Range("$G$52:$G$57"), ActiveSheet.Range("$I$45"), False, 0.05
            End If
            'Rescrever
            Range("I45").FormulaR1C1 = "Teste-F:2 amostras p/ variância"
            'Definir Tipo de Teste t
            If Application.International(xlDecimalSeparator) = "." Then
                FCalc = Range("J52").FormulaR1C1
                FCrit = Range("J54").FormulaR1C1
            Else
                FCalc = Replace(Range("J52").FormulaR1C1, ".", ",")
                FCrit = Replace(Range("J54").FormulaR1C1, ".", ",")
            End If
            If FCalc <= FCrit Then 'Mesma Variancia
                If VarVar1 >= VarVar2 Then 'Ordem Direta
                    Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$G$52:$G$57"), ActiveSheet.Range("$H$52:$H$57"), ActiveSheet.Range("$L$45"), False, 0.05
                Else 'Ordem Inversa
                    Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$H$52:$H$57"), ActiveSheet.Range("$G$52:$G$57"), ActiveSheet.Range("$L$45"), False, 0.05
                End If
                ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância equivalente"
                'Reescreve Formula
                Range("M42:O42").FormulaR1C1 = "=IF(R[12]C<>0,IF(AND(ABS(R[12]C)<=R[16]C),""PASSOU"",""FALHA""),""N/A"")"
            Else 'Diferente Variancia
                If VarVar1 >= VarVar2 Then 'Ordem Direta
                    Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$G$52:$G$57"), ActiveSheet.Range("$H$52:$H$57"), ActiveSheet.Range("$L$45"), False, 0.05
                Else 'Ordem Inversa
                    Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$H$52:$H$57"), ActiveSheet.Range("$G$52:$G$57"), ActiveSheet.Range("$L$45"), False, 0.05
                End If
                ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância diferente"
                'Reescreve Formula
                Range("M42:O42").FormulaR1C1 = "=IF(R[11]C<>0,IF(AND(ABS(R[11]C)<=R[15]C),""PASSOU"",""FALHA""),""N/A"")"
            End If
        Else
            MsgBox "Para Precisão Intermediária/Reprodutibilidade, selecione número de pontos e preencha completamente o teste!"
        End If

        'Robustez1
            If Range("B83").FormulaR1C1 = "2 Condições" Then 'IF 2 niveis
                Range("G89:H91, N83:O83").ClearContents
                'Teste F
                Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$I$89:$I$91"), ActiveSheet.Range("$J$89:$J$91"), ActiveSheet.Range("$P$92"), False, 0.05
                'Teste F usado corretamente? Variancia da primeira Variável deve ser maior do que a da segunda
                If Application.International(xlDecimalSeparator) = "." Then
                    VarVar1 = Range("Q96").FormulaR1C1
                    VarVar2 = Range("R96").FormulaR1C1
                Else
                    VarVar1 = Replace(Range("Q96").FormulaR1C1, ".", ",")
                    VarVar2 = Replace(Range("R96").FormulaR1C1, ".", ",")
                End If
                If VarVar2 > VarVar1 Then
                    'Limpar
                    Range("P92:R101").Clear
                    'Refazer o Teste F
                    Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$J$89:$J$91"), ActiveSheet.Range("$I$89:$I$91"), ActiveSheet.Range("$P$92"), False, 0.05
                End If
                'Formatação F
                Range("P92").Cut Destination:=Range("A92")
                Range("P94:R98").Cut Destination:=Range("A93:C97")
                Range("P99:R101").Cut Destination:=Range("D95:F97")
                With Range("D93:F93")
                    With .Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                    End With
                    With .Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                End With
                With Range("A97:C97").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                'Definir Tipo de Teste t
                If Application.International(xlDecimalSeparator) = "." Then
                    FCalc = Range("E95").FormulaR1C1
                    FCrit = Range("E97").FormulaR1C1
                Else
                    FCalc = Replace(Range("E95").FormulaR1C1, ".", ",")
                    FCrit = Replace(Range("E97").FormulaR1C1, ".", ",")
                End If
                If FCalc <= FCrit Then 'Mesma Variancia
                    If VarVar1 >= VarVar2 Then 'Ordem Direta
                        Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$I$89:$I$91"), ActiveSheet.Range("$J$89:$J$91"), ActiveSheet.Range("$P$92"), False, 0.05
                    Else 'Ordem Inversa
                        Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$J$89:$J$91"), ActiveSheet.Range("$I$89:$I$91"), ActiveSheet.Range("$P$92"), False, 0.05
                    End If
                    ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância equivalente"
                Else 'Diferente Variancia
                    If VarVar1 >= VarVar2 Then 'Ordem Direta
                        Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$I$89:$I$91"), ActiveSheet.Range("$J$89:$J$91"), ActiveSheet.Range("$P$92"), False, 0.05
                    Else 'Ordem Inversa
                        Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$J$89:$J$91"), ActiveSheet.Range("$I$89:$I$91"), ActiveSheet.Range("$P$92"), False, 0.05
                    End If
                    ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância diferente"
                End If
                'Formatação t
                Range("P92").Cut Destination:=Range("H92")
                Range("P94:R98").Cut Destination:=Range("H93:J97")
                Range("P99:R102").Cut Destination:=Range("K94:M97")
                Range("P103:R105").Cut Destination:=Range("M95:O97")
                With Range("H97:L97").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                With Range("K93:O93")
                    With .Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                    End With
                    With .Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                End With
                With Range("D94:D97, K94:K97, M94:M97").Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                End With
                'Reescrever formula
                 If FCalc <= FCrit Then 'Mesma Variancia
                    Range("M82:O82").FormulaR1C1 = "=IF(R[14]C[-1]<>0,IF(AND(ABS(R[14]C[-1])<=R[15]C[1]),""PASSOU"",""FALHA""),""N/A"")"
                 Else 'Diferente Variancia
                    Range("M82:O82").FormulaR1C1 = "=IF(R[13]C[-1]<>0,IF(AND(ABS(R[13]C[-1])<=R[14]C[1]),""PASSOU"",""FALHA""),""N/A"")"
                 End If
            ElseIf Range("B83").FormulaR1C1 = "3 Condições" Then 'ELSE IF 3 niveis
                'Anova
                Application.Run "ATPVBAEN.XLAM!Anova1", ActiveSheet.Range("$I$89:$K$91"), ActiveSheet.Range("$P$92"), "C", False, 0.05
                Range("P92").Cut Destination:=Range("A92")
                Range("P94:T98").Cut Destination:=Range("A93")
                Range("P101:V106").Cut Destination:=Range("G93")
                Range("G94").FormulaR1C1 = "Variação"
                Range("G95").FormulaR1C1 = "Intragrupos"
                Range("G96").FormulaR1C1 = "Intergrupos"
                'Reescrever Formula
                Range("M82:O82").FormulaR1C1 = "=IF(R[13]C[-2]<>0,IF(AND(ABS(R[13]C[-2])<=R[13]C),""PASSOU"",""FALHA""),""N/A"")"
            Else
                If Not Aviso Then
                    MsgBox "Para Robustez, selecione número de níveis e preencha completamente o teste!"
                    Aviso = True
                End If
            End If

        'Robustez2
            If Range("B101").FormulaR1C1 = "2 Condições" Then 'IF 2 niveis
                Range("G107:H109, N101:O101").ClearContents
                'Teste F
                Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$I$107:$I$109"), ActiveSheet.Range("$J$107:$J$109"), ActiveSheet.Range("$P$110"), False, 0.05
                'Teste F usado corretamente? Variancia da primeira Variável deve ser maior do que a da segunda
                If Application.International(xlDecimalSeparator) = "." Then
                    VarVar1 = Range("Q114").FormulaR1C1
                    VarVar2 = Range("R114").FormulaR1C1
                Else
                    VarVar1 = Replace(Range("Q114").FormulaR1C1, ".", ",")
                    VarVar2 = Replace(Range("R114").FormulaR1C1, ".", ",")
                End If
                If VarVar2 > VarVar1 Then
                    'Limpar
                    Range("P110:R119").Clear
                    'Refazer o Teste F
                    Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$J$107:$J$109"), ActiveSheet.Range("$I$107:$I$109"), ActiveSheet.Range("$P$110"), False, 0.05
                End If
                'Formatação F
                Range("P110").Cut Destination:=Range("A110")
                Range("P112:R116").Cut Destination:=Range("A111:C115")
                Range("P117:R119").Cut Destination:=Range("D113:F115")
                With Range("D111:F111")
                    With .Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                    End With
                    With .Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                End With
                With Range("A115:C115").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                'Definir Tipo de Teste t
                If Application.International(xlDecimalSeparator) = "." Then
                    FCalc = Range("E113").FormulaR1C1
                    FCrit = Range("E115").FormulaR1C1
                Else
                    FCalc = Replace(Range("E113").FormulaR1C1, ".", ",")
                    FCrit = Replace(Range("E115").FormulaR1C1, ".", ",")
                End If
                If FCalc <= FCrit Then 'Mesma Variancia
                    If VarVar1 >= VarVar2 Then 'Ordem Direta
                        Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$I$107:$I$109"), ActiveSheet.Range("$J$107:$J$109"), ActiveSheet.Range("$P$110"), False, 0.05
                    Else 'Ordem Inversa
                        Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$J$107:$J$109"), ActiveSheet.Range("$I$107:$I$109"), ActiveSheet.Range("$P$110"), False, 0.05
                    End If
                    ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância equivalente"
                Else 'Diferente Variancia
                    If VarVar1 >= VarVar2 Then 'Ordem Direta
                        Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$I$107:$I$109"), ActiveSheet.Range("$J$107:$J$109"), ActiveSheet.Range("$P$110"), False, 0.05
                    Else 'Ordem Inversa
                        Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$J$107:$J$109"), ActiveSheet.Range("$I$107:$I$109"), ActiveSheet.Range("$P$110"), False, 0.05
                    End If
                    ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância diferente"
                End If
                'Formatação t
                Range("P110").Cut Destination:=Range("H110")
                Range("P112:R116").Cut Destination:=Range("H111:J115")
                Range("P117:R120").Cut Destination:=Range("K112:M115")
                Range("P121:R123").Cut Destination:=Range("M113:O115")
                With Range("H115:L115").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                With Range("K111:O111")
                    With .Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                    End With
                    With .Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                End With
                With Range("D112:D115, K112:K115, M112:M115").Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                End With
                'Reescrever formula
                 If FCalc <= FCrit Then 'Mesma Variancia
                    Range("M100:O100").FormulaR1C1 = "=IF(R[14]C[-1]<>0,IF(AND(ABS(R[14]C[-1])<=R[15]C[1]),""PASSOU"",""FALHA""),""N/A"")"
                 Else 'Diferente Variancia
                    Range("M100:O100").FormulaR1C1 = "=IF(R[13]C[-1]<>0,IF(AND(ABS(R[13]C[-1])<=R[14]C[1]),""PASSOU"",""FALHA""),""N/A"")"
                 End If
            ElseIf Range("B101").FormulaR1C1 = "3 Condições" Then 'ELSE IF 3 niveis
                'Anova
                Application.Run "ATPVBAEN.XLAM!Anova1", ActiveSheet.Range("$I$107:$K$109"), ActiveSheet.Range("$P$110"), "C", False, 0.05
                Range("P110").Cut Destination:=Range("A110")
                Range("P112:T116").Cut Destination:=Range("A111")
                Range("P119:V124").Cut Destination:=Range("G111")
                Range("G112").FormulaR1C1 = "Variação"
                Range("G113").FormulaR1C1 = "Intragrupos"
                Range("G114").FormulaR1C1 = "Intergrupos"
                Range("M100:O100").FormulaR1C1 = "=IF(R[13]C[-2]<>0,IF(AND(ABS(R[13]C[-2])<=R[13]C),""PASSOU"",""FALHA""),""N/A"")"
            Else
                If Not Aviso Then
                    MsgBox "Para Robustez, selecione número de níveis e preencha completamente o teste!"
                    Aviso = True
                End If
            End If
            
        'Robustez3
            If Range("B124").FormulaR1C1 = "2 Condições" Then 'IF 2 niveis
                Range("G130:H132, N124:O124").ClearContents
                'Teste F
                Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$I$130:$I$132"), ActiveSheet.Range("$J$130:$J$132"), ActiveSheet.Range("$P$133"), False, 0.05
                'Teste F usado corretamente? Variancia da primeira Variável deve ser maior do que a da segunda
                If Application.International(xlDecimalSeparator) = "." Then
                    VarVar1 = Range("Q137").FormulaR1C1
                    VarVar2 = Range("R137").FormulaR1C1
                Else
                    VarVar1 = Replace(Range("Q137").FormulaR1C1, ".", ",")
                    VarVar2 = Replace(Range("R137").FormulaR1C1, ".", ",")
                End If
                If VarVar2 > VarVar1 Then
                    'Limpar
                    Range("P133:R142").Clear
                    'Refazer o Teste F
                    Application.Run "ATPVBAEN.XLAM!Ftestv", ActiveSheet.Range("$J$130:$J$132"), ActiveSheet.Range("$I$130:$I$132"), ActiveSheet.Range("$P$133"), False, 0.05
                End If
                'Formatação F
                Range("P133").Cut Destination:=Range("A133")
                Range("P135:R139").Cut Destination:=Range("A134:C138")
                Range("P140:R142").Cut Destination:=Range("D136:F138")
                With Range("D134:F134")
                    With .Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                    End With
                    With .Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                End With
                With Range("A138:C138").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                'Definir Tipo de Teste t
                If Application.International(xlDecimalSeparator) = "." Then
                    FCalc = Range("E136").FormulaR1C1
                    FCrit = Range("E138").FormulaR1C1
                Else
                    FCalc = Replace(Range("E136").FormulaR1C1, ".", ",")
                    FCrit = Replace(Range("E138").FormulaR1C1, ".", ",")
                End If
                If FCalc <= FCrit Then 'Mesma Variancia
                    If VarVar1 >= VarVar2 Then 'Ordem Direta
                        Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$I$130:$I$132"), ActiveSheet.Range("$J$130:$J$132"), ActiveSheet.Range("$P$133"), False, 0.05
                    Else 'Ordem Inversa
                        Application.Run "ATPVBAEN.XLAM!Ttestm", ActiveSheet.Range("$J$130:$J$132"), ActiveSheet.Range("$I$130:$I$132"), ActiveSheet.Range("$P$133"), False, 0.05
                    End If
                    ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância equivalente"
                Else 'Diferente Variancia
                    If VarVar1 >= VarVar2 Then 'Ordem Direta
                        Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$I$130:$I$132"), ActiveSheet.Range("$J$130:$J$132"), ActiveSheet.Range("$P$133"), False, 0.05
                    Else 'Ordem Inversa
                        Application.Run "ATPVBAEN.XLAM!Pttestv", ActiveSheet.Range("$J$130:$J$132"), ActiveSheet.Range("$I$130:$I$132"), ActiveSheet.Range("$P$133"), False, 0.05
                    End If
                    ActiveCell.FormulaR1C1 = "Teste-t:2 amostras:Variância diferente"
                End If
                'Formatação t
                Range("P133").Cut Destination:=Range("H133")
                Range("P135:R139").Cut Destination:=Range("H134:J138")
                Range("P140:R143").Cut Destination:=Range("K135:M138")
                Range("P144:R146").Cut Destination:=Range("M136:O138")
                With Range("H138:L138").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                With Range("K134:O134")
                    With .Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                    End With
                    With .Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                End With
                With Range("D135:D138, K135:K138, M135:M138").Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                End With
                'Reescrever formula
                 If FCalc <= FCrit Then 'Mesma Variancia
                    Range("M123:O123").FormulaR1C1 = "=IF(R[14]C[-1]<>0,IF(AND(ABS(R[14]C[-1])<=R[15]C[1]),""PASSOU"",""FALHA""),""N/A"")"
                 Else 'Diferente Variancia
                    Range("M123:O123").FormulaR1C1 = "=IF(R[13]C[-1]<>0,IF(AND(ABS(R[13]C[-1])<=R[14]C[1]),""PASSOU"",""FALHA""),""N/A"")"
                 End If
            ElseIf Range("B124").FormulaR1C1 = "3 Condições" Then 'ELSE IF 3 niveis
                'Anova
                Application.Run "ATPVBAEN.XLAM!Anova1", ActiveSheet.Range("$I$130:$K$132"), ActiveSheet.Range("$P$133"), "C", False, 0.05
                Range("P133").Cut Destination:=Range("A133")
                Range("P135:T139").Cut Destination:=Range("A134")
                Range("P142:V147").Cut Destination:=Range("G134")
                Range("G135").FormulaR1C1 = "Variação"
                Range("G136").FormulaR1C1 = "Intragrupos"
                Range("G137").FormulaR1C1 = "Intergrupos"
                Range("M123:O123").FormulaR1C1 = "=IF(R[13]C[-2]<>0,IF(AND(ABS(R[13]C[-2])<=R[13]C),""PASSOU"",""FALHA""),""N/A"")"
            Else
                If Not Aviso Then
                    MsgBox "Para Robustez, selecione número de níveis e preencha completamente o teste!"
                    Aviso = True
                End If
            End If
            
Epílogo:
            'Formatar Células Modificadas por Estatistica
            With Range("A19:F30, I45:O58, A92:O98, A110:O116, A133:O139").Font
                .Name = "Calibri": .Size = 7
                '.Strikethrough = False
                '.Superscript = False
                '.Subscript = False
                '.OutlineFont = False
                '.Shadow = False
                '.Underline = xlUnderlineStyleNone
                '.ThemeColor = xlThemeColorLight1
                '.TintAndShade = 0
                '.ThemeFont = xlThemeFontMinor
            End With
            
            ' Ajustar Página Largura Após ANOVA no FIM
                Columns("A:O").ColumnWidth = 5.86
                'Range("A1").Select
            
            'UnSpeed Up
                ActiveSheet.DisplayPageBreaks = True
                Application.EnableEvents = True
                Application.DisplayStatusBar = True
                Application.ScreenUpdating = True
                Application.Calculation = xlCalculationAutomatic
            
            'Protect
                ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="0"
                ActiveSheet.EnableSelection = xlUnlockedCells
            'Instalar Suplementos
                'AddIns("Ferramentas de Análise - VBA").Installed = False
    End If
    Exit Sub      ' Exit to avoid handler.
Erro:
        MsgBox "Foi detectado um erro na sua planilha!"
        MsgBox "Não é possível continuar. Por favor, verifique!"
        MsgBox "Em caso de Dúvida consulte a COBPF."
        GoTo Epílogo
End Sub
Private Sub Limpa1_Click()
    'UnProtect
        ActiveSheet.Unprotect Password:="0"
    
    Range("A10,C10:E10,A12:A18,C12:E18").ClearContents
    Range("A19:H33").Clear

    'Protect
            ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="0"
            ActiveSheet.EnableSelection = xlUnlockedCells
End Sub
Private Sub Limpa2_Click()
    'UnProtect
        ActiveSheet.Unprotect Password:="0"
    
    Range("A45,C45,C47:C49,G45,G47:G49,C52:F57").ClearContents
    Range("I43:O58").Clear

    'Protect
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="0"
        ActiveSheet.EnableSelection = xlUnlockedCells
End Sub
Private Sub Limpa3_Click()
    Range("C66,C68:C70,F65:G70").ClearContents
End Sub
Private Sub Limpa4_Click()
    Range("C76,C78:C80,F74:H76,E78:H80").ClearContents
End Sub
Private Sub Limpa5_Click()
    'UnProtect
        ActiveSheet.Unprotect Password:="0"

    Range("B83:C83,E83:F83,H83:I83,K83:L83,N83:O83,C85:H87,C89:H91").ClearContents
    Range("A92:O99").Clear
    
    'Protect
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="0"
        ActiveSheet.EnableSelection = xlUnlockedCells
End Sub
Private Sub Limpa6_Click()
    'UnProtect
        ActiveSheet.Unprotect Password:="0"

    Range("B101:C101,E101:F101,H101:I101,K101:L101,N101:O101,D103:H105,C107:H109").ClearContents
    Range("A110:O117").Clear

    'Protect
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="0"
        ActiveSheet.EnableSelection = xlUnlockedCells
End Sub
Private Sub Limpa7_Click()
    'UnProtect
        ActiveSheet.Unprotect Password:="0"

    Range("B124:C124,E124:F124,H124:I124,K124:L124,N124:O124,C126:H128,C130:H132").ClearContents
    Range("A133:O140").Clear

    'Protect
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="0"
        ActiveSheet.EnableSelection = xlUnlockedCells
End Sub
Private Sub LimpaTudo_Click()
    'UnProtect
        ActiveSheet.Unprotect Password:="0"
    
    Range("B5:B6,D5:D6,F5:F6,H5:H6,J5:J6,L5:L6,N5,A10,C10:E10,A12:A18,C12:E18,A45,C45,C47:C49,G45,G47:G49,C52:F57,C66,C68:C70,F65:G70,C76,C78:C80,F74:H76,E78:H80,B83:C83,E83:F83,H83:I83,K83:L83,N83:O83,C85:H87,C89:H91").ClearContents
    Range("B101:C101,E101:F101,H101:I101,K101:L101,N101:O101,D103:H105,C107:H109,B124:C124,E124:F124,H124:I124,K124:L124,N124:O124,C126:H128,C130:H132").ClearContents
    Range("A19:H33,I43:O58,A92:O99,A110:O117,A133:O140").Clear

    'Protect
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="0"
        ActiveSheet.EnableSelection = xlUnlockedCells
End Sub
