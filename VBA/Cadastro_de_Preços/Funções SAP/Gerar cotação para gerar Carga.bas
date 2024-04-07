Sub Gerar_cotacao_itens_novos()
'
' Macro2 Macro
'

'
    Dim caminho As String
    Dim fornecedor As Double, empresa As String, centro As String
    Dim carga As String
        
    caminho = Range("F3").Value
    
    Workbooks.Open (caminho)
        Range(Range("A2:I2"), Range("A2:I2").End(xlDown)).ClearContents
    'Cód Materiais
    Windows("Planilha de Duplo Check_2.xlsm").Activate
        Range(Range("T10"), Range("T10").End(xlDown)).Copy
    Windows("Template - Cotacao.xlsx").Activate
        Range("D2").Select
        ActiveSheet.Paste
    
    'Valores
    Windows("Planilha de Duplo Check_2.xlsm").Activate
        Range(Range("U10"), Range("U10").End(xlDown)).Copy
    Windows("Template - Cotacao.xlsx").Activate
        Range("H2").Select
        ActiveSheet.Paste
        
    'IVA's
    Windows("Planilha de Duplo Check_2.xlsm").Activate
        Range(Range("V10"), Range("V10").End(xlDown)).Copy
    Windows("Template - Cotacao.xlsx").Activate
        Range("I2").Select
        ActiveSheet.Paste
        
    'Fornecedores, Centro, Empresa, Qtd, Grp, e outros
    Windows("Planilha de Duplo Check_2.xlsm").Activate
        fornecedor = Range("C2").Value
        empresa = Range("C4").Value
        centro = Range("C3").Value
        
    Windows("Template - Cotacao.xlsx").Activate
    Range("C:C,F:F").NumberFormat = "@"
    If Range("D3").Value = "" Then
        Range("A2") = "1500"
        Range("B2") = "103"
        Range("C2") = empresa
        Range("E2") = "1"
        Range("F2") = centro
        Range("G2") = fornecedor
    Else
        For Each cell In Range(Range("D2"), Range("D2").End(xlDown))
            cell.Offset(0, -3) = "1500"
            cell.Offset(0, -2) = "103"
            cell.Offset(0, -1) = empresa
            cell.Offset(0, 1) = "1"
            cell.Offset(0, 2) = centro
            cell.Offset(0, 3) = fornecedor
        Next
    End If
    Windows("Template - Cotacao.xlsx").Close SaveChanges:=True
        
    Call Abrir_SAP
    
    
    With session
        
        .findById("wnd[0]/tbar[0]/okcd").Text = "zlbrr_mm_0003"
        .findById("wnd[0]").sendVKey 0

        .findById("wnd[0]/usr/ctxtP_FILE").Text = caminho
        .findById("wnd[0]").sendVKey 8
        
        .findById("wnd[0]").sendVKey 20
        
    End With
    
    carga = Right$(session.findById("wnd[0]/sbar").Text, 4)
        Application.Wait (Now + TimeValue("0:00:10"))
    
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
    
    'Cola no Excel
    Windows("Planilha de Duplo Check_2.xlsm").Activate
    Sheets("LOG").Range("B1").PasteSpecial
    Sheets("LOG").Range("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="|", TrailingMinusNumbers:=True

    'session.findById("wnd[0]/tbar[0]/btn[3]").press

    
End Sub
