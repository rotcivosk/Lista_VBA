sub criar_cotacao()

    ' Variáveis importadas da tabela
    Dim fornecedor As Double, data_i As String, data_f As String, iva As String
    fornecedor = Range("G10").Value
    data_i = Range("G8").Value
    data_f = Range("G9").Value
    iva = "S1"
    
    If Range("B22") = "" Then
        Set range_valores = Range("B21")
    Else
        Set range_valores = Range(Range("B21"), Range("B21").End(xlDown))
    End If
    range_valores.Copy

        'Variáveis Usadaas:
        'Número de Requisições
        'data_i, data_f, fornecedor, range_valores, iva, cotacao

    With session

        'Abre ME57
        .findById("wnd[0]/tbar[0]/okcd").Text = "me57"
        .findById("wnd[0]").sendVKey 0
        
        'Filtros
        
        .findById("wnd[0]/usr/btn%_BA_BANFN_%_APP_%-VALU_PUSH").press
        .findById("wnd[1]/tbar[0]/btn[24]").press
        .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]/usr/ctxtP_LSTUB").Text = "Alv"
        .findById("wnd[0]").sendVKey 8
     
        'Dentro -> RFQ sem cotação
        .findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0018/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3214/btnSELECT_ALL").press
        .findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0018/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3214/cntlMEREQ3214_CC/shellcont/shell").pressContextButton "MERFQVENDORALL"
        .findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0018/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3214/cntlMEREQ3214_CC/shellcont/shell").selectContextMenuItem "MERFQASSIGNALL"
        .findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0018/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT2").Select
        .findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0018/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT2/ssubTABSTRIPCONTROL3SUB:SAPLME57N:0002/cntlSOURCERFQ/shellcont/shell").modifyCheckbox 2, "SELKZ", True
        .findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0018/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT2/ssubTABSTRIPCONTROL3SUB:SAPLME57N:0002/cntlSOURCERFQ/shellcont/shell").currentCellRow = 2
        .findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0018/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT2/ssubTABSTRIPCONTROL3SUB:SAPLME57N:0002/cntlSOURCERFQ/shellcont/shell").pressToolbarButton "&MERFQCREATE"

        'Data de cotação
        .findById("wnd[0]/usr/ctxtEKKO-ANGDT").Text = data_i
        .findById("wnd[0]").sendVKey 0

        'Data de Remessa
        While InStr(session.findById("wnd[0]/sbar/pane[0]").Text, "encontra-se no passado") Or InStr(session.findById("wnd[0]/sbar/pane[0]").Text, "Entrar data de remessa posterior ao prazo para entrega da cotação") Or InStr(session.findById("wnd[0]/sbar/pane[0]").Text, "Data de entrega: o próximo dia útil ")
            .findById("wnd[0]/usr/ctxtRM06E-EEIND").Text = data_f
            .findById("wnd[0]").sendVKey 0
            Application.Wait Now + TimeValue("0:00:01")
        Wend

        'Fornecedor e salvar
        .findById("wnd[0]/tbar[1]/btn[7]").press
        .findById("wnd[0]/usr/ctxtEKKO-LIFNR").Text = fornecedor
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]").sendVKey 11
        
        'Abrir a ME47
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nme47"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]").sendVKey 0
    End With
    'Looping de cada valor
    Dim num_temp As Integer
    num_temp = 0
    For Each cell In range_valores
        
        If cell.Offset(-1, 0).Value <> cell.Value Or cell.Offset(-1, 1).Value <> cell.Offset(0, 1).Value Then
            With session
                'Abre a linha
                .findById("wnd[0]/usr/tblSAPMM06ETC_0323").getAbsoluteRow(ident).Selected = True
                .findById("wnd[0]").sendVKey 16
                
                'IVA
                .findById("wnd[0]/usr/ctxtEKPO-MWSKZ").Text = iva
                .findById("wnd[0]").sendVKey 0
            End With
        End If
        
        
        ' Se a linha de baixo for igual
        If cell.Offset(1, 0).Value = cell.Value And cell.Offset(1, 1).Value = cell.Offset(0, 1).Value Then
            With session
                ' Vou colocar apenas o valor desta linha
                .findById("wnd[0]/usr/subSERVICE:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/txtESLL-TBTWR[6," & num_temp & "]").Text = cell.Offset(0, 6).Value
                .findById("wnd[0]").sendVKey 0
            End With
            num_temp = num_temp + 1
        Else
            With session
                'Vou colocar o valor desta linha
    
                'Valor Final
                .findById("wnd[0]/usr/subSERVICE:SAPLMLSP:0400/tblSAPLMLSPTC_VIEW/txtESLL-TBTWR[6," & num_temp & "]").Text = cell.Offset(0, 6).Value
                .findById("wnd[0]").sendVKey 0
                
                If cell.Offset(0, 5) <> "" Then
                    'Valor Inicial
                    .findById("wnd[0]/usr/subSERVICE:SAPLMLSP:0400/btnCONDITION").press
                    .findById("wnd[0]/usr/subCONDITIONS:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,9]").Text = "ZPBI"
                    .findById("wnd[0]/usr/subCONDITIONS:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,9]").Text = cell.Offset(0, 5).Value
                    .findById("wnd[0]").sendVKey 0
                    .findById("wnd[0]").sendVKey 3
                End If
                .findById("wnd[0]").sendVKey 3
            End With
            num_temp = 0
            ident = ident + 1
        End If
    Next
        
    'Salvar
    session.findById("wnd[0]").sendVKey 11
    
    'Número da Cotação
    Range("G6").Value = session.findById("wnd[0]/usr/ctxtRM06E-ANFNR").Text
    



end sub