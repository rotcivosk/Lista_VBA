Public session

Sub emitir_pedidos()

    '***_____BASE____****
   
   
   'Abrir o SAP
    Dim Applic, Connection, SapGuiAuto
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Applic = SapGuiAuto.GetScriptingEngine
    Set Connection = Applic.Children(0)
    Set session = Connection.Children(0) 'Declara a Session pública como o SAP em aberto

    'Variáveis importadas da tabela
    Dim fornecedor As Double, requisicao As Double, cotacao As Double
    Dim data_i As String, data_f As String, iva As String, texto_padrao As String
    requisicao = Range("d2").Value
    fornecedor = Range("G10").Value
    data_i = Range("G8").Value
    data_f = Range("G9").Value
    texto_padrao = Range("D46").Value
    iva = "S1"
    
    
    'Variáveis de Anexos
    Dim caminho_vbs As String, caminho_anexos As String
    caminho_anexos = "D:\Users\sb048948\OneDrive - Honda\Documentos\SAP\SAP GUI\"
    caminho_vbs = "D:\Users\sb048948\Downloads\Emitir_pedidos\"
    
    'Nomes
    Dim proposta As String
    proposta = Range("d41").Value

    'Seleciona o range que vai trabalhar e o range de anexos
    If Range("F41") = "" Then
        Set range_anexos = Range("F40")
    Else
        Set range_anexos = Range(Range("F40"), Range("F40").End(xlDown))
    End If
    
    
    
    If Range("B22") = "" Then
        Set range_valores = Range("B21")
    Else
        Set range_valores = Range(Range("B21"), Range("B21").End(xlDown))
    End If
    range_valores.Copy
    'Identador
    Dim ident As Integer
    ident = 0







    '***__CRIAR_COTAÇÃO_****

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
    cotacao = session.findById("wnd[0]/usr/ctxtRM06E-ANFNR").Text
    Range("G6").Value = cotacao
    




    '***____CRIAR_PEDIDO____***
    
    
    
    'Abrir ME21N
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme21n"
    session.findById("wnd[0]").sendVKey 0
    
    
    
    
    'Transferir a cotação
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    On Error GoTo 0
    With session
    'Abre o menu de cotação

        .findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").pressContextButton "SELECT"
        .findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").selectContextMenuItemByPosition 5
    
    'Limpa o campo de cotação
        .findById("wnd[0]/usr/txtP_QCOUNT").Text = ""
        .findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00023-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00022-LOW").Text = ""
        .findById("wnd[0]/usr/txtSP$00021-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00024-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00018-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00019-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00015-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00014-LOW").Text = cotacao
        .findById("wnd[0]/usr/txtSP$00026-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00012-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00013-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00011-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtSP$00017-LOW").Text = ""
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
       
    
End Sub
Sub Adicionar_anexos_da_lista()
    
    'Variáveis de Anexos
    Dim caminho_vbs As String, caminho_anexos As String
    caminho_anexos = "D:\Users\sb048948\OneDrive - Honda\Documentos\SAP\SAP GUI\"
    caminho_vbs = "D:\Users\sb048948\Downloads\Emitir_pedidos\"

    'Abrir o SAP
    Dim Applic, Connection, SapGuiAuto
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Applic = SapGuiAuto.GetScriptingEngine
    Set Connection = Applic.Children(0)
    Set session = Connection.Children(0) 'Declara a Session pública como o SAP em aberto

    'Seleciona o range que vai trabalhar e o range de anexos
    If Range("F41") = "" Then
        Set range_anexos = Range("F40")
    Else
        Set range_anexos = Range(Range("F40"), Range("F40").End(xlDown))
    End If

    Dim temp As Boolean
    

    '***___CRIAR_ANEXOS___***
    For Each cell In range_anexos
        If cell.Offset(0, 1).Value = "Contrato" Then temp = True Else temp = False
        Call adicionar_anexos(caminho_anexos, caminho_vbs, cell.Value, temp, False)
    Next

End Sub

Function checar_saplmegui()
    
    Dim nome_comp As String
            
    Set user_megui = session.findById("wnd[0]/usr")
    i = 0
    For i = 0 To user_megui.Children.Count - 1
        nome_comp = user_megui.Children(CInt(i)).Name
        If Left(nome_comp, 15) = "SUB0:SAPLMEGUI:" Then
            Exit For
        End If
    Next
    
    checar_saplmegui = Right(nome_comp, 4)
    
End Function

