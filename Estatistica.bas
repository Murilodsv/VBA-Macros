Attribute VB_Name = "Módulo1"
Sub ESTATISTICA()
'
' ESTATISTICA MANIPULA A PLANILHA COM VALORES ESTATISTICOS PARA COMPARACAO DE MODELOS
'

'

Calculate


n_est = Sheets("BASE_ESTAT").Range("R" & 1).Value
N = Sheets("BASE_ESTAT").Range("R" & 2).Value
unidade = Sheets("ENTRADA").Range("J" & 2).Value
NLINHA = 4

For X = 1 To n_est

If X = 1 Then

Sheets("SAIDA").Select
Columns("AB:AM").Select
Selection.Cut
Sheets.Add After:=Sheets(Sheets.Count)
Sheets((Sheets.Count)).Select
Columns("A:A").Select
ActiveSheet.Paste
Sheets((Sheets.Count)).Select
Application.DisplayAlerts = False
ActiveWindow.SelectedSheets.Delete

Calculate

Sheets("SAIDA").Select
Sheets("BASE_ESTAT").Visible = True

Sheets("BASE_ESTAT").Select
Range("C7:M7").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Range("C6").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy

Range(Cells(6, 3), Cells(N + 5, 3)).Select
ActiveSheet.Paste

Range("A6:B6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Sheets("SAIDA").Select
Range("A3:B3").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Calculate

Sheets("ENTRADA").Select
Range(Cells(5, 2), Cells(5, n_est + 1)).Select
Selection.Copy

Sheets("SAIDA").Select
Range("A3").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

End If
Calculate
Sheets("ENTRADA").Select
Range(Cells(6, 1), Cells(N + 5, 1)).Select
Selection.Copy

Sheets("BASE_ESTAT").Select
Range("A6").Select
ActiveSheet.Paste

ActiveSheet.ChartObjects("Gráfico 1").Activate
ActiveChart.SeriesCollection(1).XValues = "=BASE_ESTAT!$A$6:$A$" & 5 + N & ""
ActiveChart.SeriesCollection(1).Values = "=BASE_ESTAT!$B$6:$B$" & 5 + N & ""
Calculate
Sheets("ENTRADA").Select
titulo = Range(Cells(5, X + 1), Cells(5, X + 1)).Value
Range(Cells(6, X + 1), Cells(N + 5, X + 1)).Select
Selection.Copy

Sheets("BASE_ESTAT").Select
Range("B6").Select
ActiveSheet.Paste

Calculate

Range("R5").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets("SAIDA").Select
Range(Cells(X + 2, 2), Cells(X + 2, 2)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True

Calculate
'---------GRÁFICO-----------

maximo = Sheets("BASE_ESTAT").Range("AG" & 2).Value
minimo = Sheets("BASE_ESTAT").Range("AG" & 3).Value

    Sheets("BASE_ESTAT").Select
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.SeriesCollection(1).Trendlines(1).Select
    Selection.Delete
    ActiveChart.Axes(xlCategory).MaximumScale = maximo
    ActiveChart.Axes(xlValue).MaximumScale = maximo
    ActiveChart.Axes(xlCategory).MinimumScale = minimo
    ActiveChart.Axes(xlValue).MinimumScale = minimo
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).Trendlines.Add
    ActiveChart.SeriesCollection(1).Trendlines(1).Select
    Selection.DisplayEquation = True
    Selection.DisplayRSquared = True
    ActiveSheet.ChartObjects("Gráfico 1").Activate
    ActiveChart.SeriesCollection(1).Trendlines(1).DataLabel.Select
    Selection.Left = 40
    Selection.Top = 30
    
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Observado " & "" & unidade & ""

    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Estimado " & "" & unidade & ""
    
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "" & titulo & ""

    ActiveSheet.Shapes.Range(Array("Gráfico 1")).Select
    ActiveChart.ChartArea.Copy
    
    Sheets("SAIDA").Select
    Range(Cells(NLINHA, 30), Cells(NLINHA, 30)).Select
    ActiveSheet.Pictures.Paste.Select
    
    Range(Cells(NLINHA - 1, 30), Cells(NLINHA - 1, 30)).Select
    ActiveCell.FormulaR1C1 = titulo


    NLINHA = NLINHA + 17
    
'---------------------------

Next
Sheets("BASE_ESTAT").Visible = True
Sheets("BASE_ESTAT").Select
ActiveWindow.SelectedSheets.Visible = False
Sheets("SAIDA").Select
Range("A1").Select
    
End Sub

Sub LIMPA()
'
' LIMPA OS VALORES DA PLANILHA
'

'

Sheets("SAIDA").Select
Columns("AB:AM").Select
Selection.Cut
Sheets.Add After:=Sheets(Sheets.Count)
Sheets((Sheets.Count)).Select
Columns("A:A").Select
ActiveSheet.Paste
Sheets((Sheets.Count)).Select
Application.DisplayAlerts = False
ActiveWindow.SelectedSheets.Delete

Sheets("ENTRADA").Select
Range("A6:B6").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

n_est = Sheets("BASE_ESTAT").Range("R" & 1).Value
N = Sheets("BASE_ESTAT").Range("R" & 2).Value

Sheets("SAIDA").Select
Sheets("BASE_ESTAT").Visible = True

Sheets("BASE_ESTAT").Select
Range("C7:M7").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Range("A6:B6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Sheets("SAIDA").Select
Range("A3:B3").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents
    
Sheets("BASE_ESTAT").Select
ActiveWindow.SelectedSheets.Visible = False
Sheets("SAIDA").Select

Sheets("ENTRADA").Select
Range("B5").Select
ActiveCell.FormulaR1C1 = "Modelo 1"

Range("C5").Select
ActiveCell.FormulaR1C1 = "Modelo 2"

Range("D5").Select
ActiveCell.FormulaR1C1 = "..."

Range("A6").Select
    
End Sub

Sub SAIR()
'
' LIMPA E FECHA A PLANILHA
'

'

Call LIMPA

Application.DisplayAlerts = False
ActiveWindow.Close

End Sub

