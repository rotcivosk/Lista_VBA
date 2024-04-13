Attribute VB_Name = "M_XLS_Importar_AP"
Sub Importar_AP()
'
' Importar_AP Macro
'

'
    ' Apagar da Tela Principal
    Range("A15").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).ClearContents
    
    ' Copiar da AP
    Sheets("AP").Select
    Range("B29:U29").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    ' Colar na Tela Principal
    Sheets("Tela Principal").Select
    Range("A14").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Formatação
    Range("A14:BE14").Copy
    Range("A15:BE15").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("U14:BE14").Copy

    Range("T15").Select
    Selection.End(xlDown).Select
    Selection.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    
End Sub


