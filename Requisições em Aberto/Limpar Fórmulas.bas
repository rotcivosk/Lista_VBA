Sub Limpar_Fórmulas()
'
' Limpar_Fórmulas Macro
'

'
    Sheets("PED - SAP").Select
    ActiveWorkbook.Worksheets("PED - SAP").AutoFilter.Sort.SortFields.Clear
    ActiveSheet.ShowAllData
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("PED - SAP").Select
    ActiveWindow.SelectedSheets.Visible = False
    Cells.Select
    Range("K1").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("REQ - SAP").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("F - Temp").Select
    ActiveWindow.SelectedSheets.Visible = False
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("F - EKKO").Select
    ActiveWindow.SelectedSheets.Visible = False
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Relato Semana Anterior").Select
    ActiveWindow.SelectedSheets.Visible = False
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("COT - ANTERIOR").Select
    ActiveWindow.SelectedSheets.Visible = False
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("F - APROV").Select
    Application.CutCopyMode = False
    Selection.Copy
    Cells.Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("F - APROV").Select
    ActiveWindow.SelectedSheets.Visible = False
    Cells.Select
    Range("A223").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("REQ - JDE").Select
    ActiveWindow.SelectedSheets.Visible = False
    Cells.Select
    Range("B1").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("PED - JDE").Select
    ActiveWindow.SelectedSheets.Visible = False
    Cells.Select
    Range("O1").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Ped - Consolidado").Select
    Cells.Select
    Range("P1").Activate
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    ActiveSheet.Previous.Select
    Range("A1").Select
    ActiveSheet.Previous.Select
    Range("A1").Select
    ActiveSheet.Previous.Select
    Range("A1:J2").Select
    ActiveSheet.Previous.Select
    Range("D1").Select
    ActiveSheet.Previous.Select
    Range("A1").Select
    ActiveSheet.Previous.Select
    Range("A1").Select
End Sub
