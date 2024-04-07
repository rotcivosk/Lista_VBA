Sub exportar_clipboardSAP(Is_tabela As Boolean)

    'Há duas maneiras de Exportar, uma caso seja uma tabela, outra caso seja uma transação
    With session
        If Is_tabela Then
            .findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
            .findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem "&PC"
        Else
            .findById("wnd[0]/tbar[1]/btn[45]").press
        End If
        .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
        .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
        .findById("wnd[1]/tbar[0]/btn[0]").press

    End With
    
End Sub

Sub export_Formatar_PlanilhaSAP()
        
        
        Range(Range("A1"), ActiveCell.SpecialCells(xlLastCell)).ClearContents
        Range("A1").PasteSpecial
        Application.CutCopyMode = False
        With Columns("A:A")
            .TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="|", TrailingMinusNumbers:=True
            .Delete Shift:=xlToLeft
        
        Rows("1:3").Delete Shift:=xlUp
        Rows("2:2").Delete Shift:=xlUp
        End With
        
        Columns("A:A").Select
        With Range(Selection, Selection.End(xlToRight))
            .EntireColumn.AutoFit
            .NumberFormat = "General"
        End With
        
End Sub
