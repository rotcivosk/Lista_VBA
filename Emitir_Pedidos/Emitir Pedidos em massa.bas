Sub Rodar_macro()
'
'
    'Range(Range("F40"), Range("F40").End(xlDown))
    'Range ("A1:O" & Range("A" & Rows.Count).End(xlUp).Row)
   
    
    Dim i As Integer, total As Integer, lin As Integer

    Sheets("base de informações").Activate
    
    total = Range("B2").Value
    
    For i = 1 To total

        Call Limpar_Campos
        
        Sheets("Base de Informações").Activate
        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0
        With Range("A1:O" & Range("A" & Rows.Count).End(xlUp).Row)
            .AutoFilter Field:=1, Criteria1:=i
            .Copy
        End With
    
        Sheets("Temp").Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Range(Range("A1"), Range("A1").End(xlDown)).Select
        lin = Selection.Rows.Count - 1
        
        If lin > 1 Then
            
            ActiveWorkbook.Worksheets("Temp").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Temp").Sort.SortFields.Add2 Key:=Range("B2:B" & lin), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With ActiveWorkbook.Worksheets("Temp").Sort
                .SetRange Range("B2:H" & lin)
                .Header = xlGuess
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        
            Range("B2:H" & Range("B" & Rows.Count).End(xlUp).Row).Copy

        Else
            Range("B2:H2").Copy
        End If
    
        
        
        
        Sheets("Macro - Pedidos").Select
        Range("B21").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
            
        Sheets("Temp").Select
        
        
        If lin > 1 Then
            ActiveWorkbook.Worksheets("Temp").Sort.SortFields.Clear
            ActiveWorkbook.Worksheets("Temp").Sort.SortFields.Add2 Key:=Range("N2:N" & lin), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With ActiveWorkbook.Worksheets("Temp").Sort
                .SetRange Range("L2:N" & lin)
                .Header = xlGuess
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            Range("L2:N" & Range("L" & Rows.Count).End(xlUp).Row).Copy
        
        Else
            Range("L2:N2").Copy
        End If
    
            
        
        Sheets("Macro - Pedidos").Select
        Range("F40").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
        DoEvents
        Call emitir_pedidos
        
    
    Next
        

End Sub