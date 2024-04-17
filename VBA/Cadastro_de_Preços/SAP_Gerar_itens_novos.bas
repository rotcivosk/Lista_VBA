Attribute VB_Name = "M_SAP_Gerar_itens_novos"
Sub Gerar_cotacao_itens_novos()
    
    Dim fornecedor As String, empresa As String, centro As String
        fornecedor = Range("M6").Value
        centro = Range("M7").Value
        If centro = "0212" Then empresa = "0200" Else empresa = "0300"
        Dim carga As String




    Dim lista As Range
    Dim item_ini As Range
        Set item_ini = Range("K10")
        If item_ini.Offset(1, 0) = "" Then
            Set lista = item_ini
        Else
            Set lista = Range(item_ini, item_ini.End(xlDown))
        End If
     Dim output() As Variant
     ReDim output(1 To 9, 1 To lista.Count)
    ' DeclaraÃ§Ãµes de VarÃ­aveis

    ' Selecionar a lista
    
    ThisWorkbook.Activate
    i = 1
    For Each Cell In lista
        output(1, i) = "1500"
        output(2, i) = "103"
        output(3, i) = empresa
        output(4, i) = Cell.Value
        output(5, i) = "1"
        output(6, i) = centro
        output(7, i) = fornecedor
        output(8, i) = Round(Cell.Offset(0, 1).Value, 2)
        output(9, i) = Cell.Offset(0, 2).Value
        i = i + 1
        Next

    ' Importar arquivos da planilha para uma matriz

    Dim caminho As String
        caminho = Range("L8").Value
        Workbooks.Open (caminho)
        Range(Range("A2:I2"), Range("A2:I2").End(xlDown)).ClearContents
        Range("A2").Resize(UBound(output, 2), UBound(output, 1)).Value = Application.Transpose(output)
        Windows("Template - Cotacao.xlsx").Close SaveChanges:=True
    ' Colar Matriz no workbook
    Call Abrir_SAP
    
    
    With session
        
        .findById("wnd[0]/tbar[0]/okcd").Text = "zlbrr_mm_0003"
        .findById("wnd[0]").sendVKey 0

        .findById("wnd[0]/usr/ctxtP_FILE").Text = caminho
        .findById("wnd[0]").sendVKey 8
        
        .findById("wnd[0]").sendVKey 20
        End With
        carga = Right$(session.findById("wnd[0]/sbar").Text, 4)
    Application.Wait (Now + TimeValue("0:00:05"))
    
    With session
        'Abrir a carga
        .findById("wnd[0]").sendVKey 25
        .findById("wnd[0]/usr/ctxtS_PROC-LOW").Text = carga
        .findById("wnd[0]").sendVKey 8
        'Exportar para a clipboard
        .findById("wnd[0]/tbar[1]/btn[45]").press
            .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
            .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
            .findById("wnd[1]/tbar[0]/btn[0]").press
        End With
        Application.Wait (Now + TimeValue("0:00:02"))
    ' Carregar a cotaÃ§Ã£o
    
    ThisWorkbook.Activate
    Sheets("LOG").Range("B1").PasteSpecial
    Sheets("LOG").Range("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="|", TrailingMinusNumbers:=True
    'session.findById("wnd[0]/tbar[0]/btn[3]").press
    
End Sub


Sub Gerar_carga_itens_novos()
'
    Dim n As Integer

    Dim lista As Range, item_ini As Range
        Set item_ini = Range("O10")
        If item_ini.Offset(1, 0) = "" Then
            Set lista = item_ini
        Else
            Set lista = Range(item_ini, item_ini.End(xlDown))
        End If
        ' Selecionar a lista


    Call Abrir_SAP
    
    'Abrir o ZI9
    session.findById("wnd[0]/tbar[0]/okcd").Text = "zi9_mm_reginfo"
    session.findById("wnd[0]").sendVKey 0
    
    For Each Cell In lista
        If Cell.Offset(0, 1) = "" Then
        With session
            'Add o 1500, o fornecedor e o Centro 0212
            .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC2").Select
            .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC2/ssubTBS_100_SCA:ZI9_MM_REGINFO:0102/subSBS_0105:ZI9_MM_REGINFO:0105/ctxtS_EBELN-LOW").Text = Cell.Value
            .findById("wnd[0]").sendVKey 8
            
            'Rodar o relatÃ³rio
            .findById("wnd[0]/usr/txtCPO_TEXT").Text = Range("H8").Value
            .findById("wnd[0]/tbar[1]/btn[8]").press
            End With
        Cell.Offset(0, 1).Formula = Right$(Left$(session.findById("wnd[0]/sbar").Text, 31), 4)
          End If
        Next
    
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    End Sub
