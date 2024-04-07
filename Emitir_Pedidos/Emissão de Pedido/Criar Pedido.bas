sub criar_pedido()

        '***____CRIAR_PEDIDO____***
    
    dim texto_padrao As String, cotacao as Double
    texto_padrao = Range("D46").Value
    cotacao = Range("G6").Value

    ' Abrir transação ME21N
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme21n"
    session.findById("wnd[0]").sendVKey 0
        
    ' Abrir "MENU DE ANEXOS" caso esteja fechado
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    On Error GoTo 0

    With session
    ' Abre o menu de cotação

        .findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").pressContextButton "SELECT"
        .findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").selectContextMenuItemByPosition 5
    
    ' Limpa o campos que não serão utilizados
        .findById("wnd[0]/usr/txtP_QCOUNT").Text = ""
        .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00023-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00022-LOW").Text = ""
        .findById("wnd[0]/usr/txtSP$00021-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00024-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00018-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00019-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00015-LOW").Text = ""
        .findById("wnd[0]/usr/txtSP$00026-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00012-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00013-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00011-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00017-LOW").Text = ""

        ' Adiciona Informações
        .findById("wnd[0]/usr/ctxtSP$00014-LOW").Text = cotacao
        .findById("wnd[0]").sendVKey 8
    End With
    
    'Clicar botão
    Call clicar_botao_cotacao
    
    
    
    '****Criar um looping****
    
    For a = 0 To range_valores.Count - 1
        'Abre a aba "Fornecimento" e flaga o "REVFATEM"
        Dim text_temp As String
        text_temp = checar_saplmegui
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & text_temp & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT6").Select
        text_temp = checar_saplmegui
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & text_temp & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1313/chkMEPO1313-WEUNB").Selected = True
        
        'Clica no próximo botao
        text_temp = checar_saplmegui
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & text_temp & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press
    Next
    
    'Textos
    With session
        'Adiciona o texto padrão
        text_temp = checar_saplmegui
        .findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & text_temp & "/subSUB1:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4000/btnDYN_4000-BUTTON").press
        text_temp = checar_saplmegui
        .findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & text_temp & "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3").Select
        text_temp = checar_saplmegui
        .findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & text_temp & "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").Text = texto_padrao
    
    End With
    
    '***___CRIAR_ANEXOS___***
    Call Adicionar_anexos_da_lista
    
    
    '***____Salvar____****
    With session
        .findById("wnd[0]/tbar[0]/btn[11]").press
        .findById("wnd[0]/tbar[1]/btn[17]").press
        .findById("wnd[1]").sendVKey 0
    End With
    
    Range("G7").Value = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/txtMEPO_TOPLINE-EBELN").Text
    
    session.findById("wnd[0]").sendVKey 3
       

end sub