Attribute VB_Name = "Módulo1"
Sub criaWTH()
'
' Macro23 Macro
'

'

'Workbooks.Open Filename:="C:\Murilo\MACRO\WTH_DSSAT.xlsx"
Application.Calculation = xlManual
dire = ThisWorkbook.Path

For X = 1 To 1

Windows("WTH_DSSAT.xlsm").Activate
Sheets("WTH_FINAL").Select

Calculate

Range("A6:A6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets("LISTA").Select
Range("A1:A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
ActiveSheet.Range("$A$1:$A$12054").RemoveDuplicates Columns:=1, Header:= _
        xlNo
        
Calculate

nano = Sheets("LISTA").Range("C" & 1).Value - 2
WTH = Sheets("ENTRADA").Range("B" & 4).Value

For y = 1 To nano

ano = Sheets("LISTA").Range("A" & y + 1).Value

Sheets("EXPORTA").Select
Range("A6:A400").Select
Selection.ClearContents

Sheets("WTH_FINAL").Select
ActiveSheet.Range("$A$5:$A$12058").AutoFilter Field:=1, Criteria1:=ano

Range("U5").Select
Selection.End(xlDown).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets("EXPORTA").Select
Range("A6").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Columns("A:A").Select
Selection.Copy

Workbooks.Add
Columns("A:A").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Application.DisplayAlerts = False
ActiveWorkbook.SaveAs Filename:="" & dire & "\" & WTH & "" & ano & "01.WTH", _
       FileFormat:=xlTextPrinter, CreateBackup:=False

Application.DisplayAlerts = False
ActiveWindow.Close
 
 
Next

Sheets("WTH_FINAL").Select
ActiveSheet.Range("$A$5:$A$12058").AutoFilter Field:=1
Sheets("ENTRADA").Select

Next

End Sub


Sub LIMPAR()
'
' Macro23 Macro
'

'

'Workbooks.Open Filename:="C:\Murilo\MACRO\WTH_DSSAT.xlsx"

Windows("WTH_DSSAT.xlsm").Activate
Sheets("ENTRADA").Select

Range("B1:B4").Select
Selection.ClearContents

Range("A7:J7").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Sheets("EXPORTA").Select
Range("A12:A13000").Select
Selection.ClearContents
Sheets("ENTRADA").Select

End Sub




