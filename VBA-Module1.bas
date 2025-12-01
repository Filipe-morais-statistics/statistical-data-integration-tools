Attribute VB_Name = "Module1"


Sub ficheiro1()

Dim ficheiro1 As String
Dim ficheiro As String

Application.ScreenUpdating = False

   Sheets("VAL_S13_C").Select
    'ficheiro 1
    ficheiro1 = "G:\aaaa\FicheirosRIE\3. carregaRIE_qfagg_valores_S13.xlsx"
    Workbooks.Open Filename:=ficheiro1
    ficheiro = ActiveWindow.Caption
    Windows(ficheiro).Activate
    
    Range("A2:C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
    Range("A2").Select
    
    Range("d5:f5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
    Range("A2").Select
    
    Windows("1. QFAGG.xlsm").Activate
    Sheets("VAL_S13_C").Select

Dim Sh As Worksheet
For Each Sh In Sheets(Array("VAL_S13_C", "VAL_S13_N"))
        Sheets("Per�odos a exportar").Select
        Columns("B:B").Select
        Selection.Copy

        Sh.Select
        Range("A1").Select
        ActiveSheet.Paste
               
        Range("A45").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(0, 0).Range("A1").Select
        linha = ActiveCell.Row
        ActiveCell.Offset(0, 0).Range("A1").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(0, 0).Range("A1").Select
        If Right(ActiveCell.Row, 5) > 9999 Then
                linhafim = linha
        Else
                linhafim = ActiveCell.Row
        End If
                linha1 = linha & ":" & linha
    
  
       Range("C32").Select
       Range(Selection, Selection.End(xlToRight)).Select
       Selection.Copy
    
    Windows(ficheiro).Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Windows("1. QFAGG.xlsm").Activate
    ActiveWindow.LargeScroll ToRight:=-1
    
    Range("B" & linha).Select
    'periodo
    Selection.Copy
    Windows(ficheiro).Activate
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select

    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    'fim periodo
    Windows("1. QFAGG.xlsm").Activate
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows(ficheiro).Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Windows("1. QFAGG.xlsm").Activate
    
    celulaatual = ActiveCell.Offset(1, -1).Range("A1").Address 
                If ActiveCell.Offset(1, -1).Range("A1").Value = vbNullString Then
                ActiveCell.Offset(1, 0).Range("A1").Select
                End If

    Application.CutCopyMode = False

    
While Not ActiveCell.Value = vbNullString
    
    Range("C32").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Windows(ficheiro).Activate
    
    ActiveCell.Offset(0, -2).Range("A1:B1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Windows("1. QFAGG.xlsm").Activate
    Range(celulaatual).Select
    
    
    Selection.Copy
    'periodo
    Windows(ficheiro).Activate
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select

    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    
    'fim periodo
    
    Windows(ficheiro).Activate
    
    Windows("1. QFAGG.xlsm").Activate
   
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows(ficheiro).Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    Windows("1. QFAGG.xlsm").Activate
    ActiveCell.Offset(1, -1).Range("A1").Select
    celulaatual = 0
    celulaatual = ActiveCell.Address
    Application.CutCopyMode = False
Wend

    Windows(ficheiro).Activate
    ActiveCell.Offset(0, -2).Range("A1:B1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select

    Windows("1. QFAGG.xlsm").Activate
Next

Application.ScreenUpdating = True
Windows(ficheiro).Activate
    Range("D4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    ActiveWorkbook.Close SaveChanges:=True

Windows("1. QFAGG.xlsm").Activate
Worksheets("Menu").Activate
Range("M23") = Now()

'MsgBox ("RIE completo S13")

End Sub
