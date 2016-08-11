Attribute VB_Name = "Módulo1"
Sub criaMET()
'
' Macro23 Macro
'

'

'Workbooks.Open Filename:="C:\Murilo\MACRO\WTH_DSSAT.xlsx"
Application.Calculation = xlManual
dire = ThisWorkbook.Path

For X = 1 To 1

Windows("MET_APSIM.xlsm").Activate
Sheets("MET_FINAL").Select

Calculate

MET = Sheets("ENTRADA").Range("B" & 3).Value

Sheets("EXPORTA").Select
Range("A12:A13000").Select
Selection.ClearContents

Sheets("MET_FINAL").Select
ActiveSheet.Range("$A$5:$A$12058").AutoFilter Field:=1, Criteria1:="<>"

Range("AA5").Select
Selection.End(xlDown).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets("EXPORTA").Select
Range("A12").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Columns("A:A").Select
Selection.Copy

Sheets("ENTRADA").Select

Workbooks.Add
Columns("A:A").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Application.DisplayAlerts = False
ActiveWorkbook.SaveAs Filename:="" & dire & "\" & MET & ".MET", _
       FileFormat:=xlTextPrinter, CreateBackup:=False

Application.DisplayAlerts = False
ActiveWindow.Close

Next

End Sub


Sub LIMPAR()
'
' Macro23 Macro
'

'

'Workbooks.Open Filename:="C:\Murilo\MACRO\WTH_DSSAT.xlsx"

Windows("MET_APSIM.xlsm").Activate
Sheets("ENTRADA").Select

Range("B1:B5").Select
Selection.ClearContents

Range("A8:L8").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Sheets("EXPORTA").Select
Range("A12:A13000").Select
Selection.ClearContents
Sheets("ENTRADA").Select

End Sub




