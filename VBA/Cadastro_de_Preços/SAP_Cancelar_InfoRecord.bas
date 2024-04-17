Attribute VB_Name = "M_SAP_Cancelar_InfoRecord"
Sub Cancelar_InfoRecord()
    '
    ' Cancelamento de Inforecord nas transacoes ME15 e ME01
    '
    Call Abrir_SAP

    ' Selecionar a lista a ser cancelada
        Dim lista As Range, item_ini As Range, temp As Integer
        Set item_ini = Range("B10")
        If item_ini.Offset(1, 0) = "" Then
            Set lista = item_ini
        Else
            Set lista = Range(item_ini, item_ini.End(xlDown))
        End If

    ' Abrir TransaÃ§Ã£o ME15
    session.findById("wnd[0]/tbar[0]/okcd").Text = "me15"
    session.findById("wnd[0]").sendVKey 0

    For Each Cell In lista
        If Cell.Offset(0, 3).Value = "" Then
        Select Case Cell.Offset(0, 2)
            Case "AMBOS"
                temp = alterar_ME15(Cell.Value, Cell.Offset(0, 1).Value, "0212", True)
                temp = alterar_ME15(Cell.Value, Cell.Offset(0, 1).Value, "0304", True)
            Case "HDA"
                temp = alterar_ME15(Cell.Value, Cell.Offset(0, 1).Value, "0212", True)
            Case "HCA"
                temp = alterar_ME15(Cell.Value, Cell.Offset(0, 1).Value, "0304", True)
            Case Else
                msgbox ("Item " & Cell.Value & " nao foi selecionada a filial para ser cancelada")
            End Select
           If temp = 0 Then Cell.Offset(0, 3).Value = "Nao ha reginfo no centro"
             End If
        Next

    ' Abrir transacao ME01
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme01"
    session.findById("wnd[0]").sendVKey 0

    For Each Cell In lista
        If Cell.Offset(0, 3).Value = "" Then
        Select Case Cell.Offset(0, 2)
            Case "AMBOS"
                temp = cancelar_me01(Cell.Value, Cell.Offset(0, 1).Value, "0212")
                temp = cancelar_me01(Cell.Value, Cell.Offset(0, 1).Value, "0304")
            Case "HDA"
                temp = cancelar_me01(Cell.Value, Cell.Offset(0, 1).Value, "0212")
            Case "HCA"
                temp = cancelar_me01(Cell.Value, Cell.Offset(0, 1).Value, "0304")
            Case Else
                msgbox ("Item " & Cell.Value & " nao foi selecionada a filial para ser cancelada")
            End Select
            If temp = 1 Then Cell.Offset(0, 3).Value = "Cancelado"
               If temp = 0 Then Cell.Offset(0, 3).Value = "Mat Bloqueado"
     End If
     Next

    session.findById("wnd[0]").sendVKey 3
    End Sub

Sub Descancelar_RegInfo()
    '
    ' Cancelamento de Inforecord nas transacoes ME15 e ME01
    '
    Call Abrir_SAP

    ' Selecionar a lista a ser cancelada
    
          Dim temp As Integer
        Dim lista As Range, item_ini As Range
        Set item_ini = Range("B10")
        If item_ini.Offset(1, 0) = "" Then
            Set lista = item_ini
        Else
            Set lista = Range(item_ini, item_ini.End(xlDown))
        End If

        ' Abrir a transaÃ§Ã£o ME15
        session.findById("wnd[0]/tbar[0]/okcd").Text = "me15"
        session.findById("wnd[0]").sendVKey 0
        
    For Each Cell In lista
    If Cell.Offset(0, 3).Value = "" Then
        Select Case Cell.Offset(0, 2)
            Case "AMBOS"
                temp = alterar_ME15(Cell.Value, Cell.Offset(0, 1).Value, "0212", False)
                If temp = 1 Then
                temp = alterar_ME15(Cell.Value, Cell.Offset(0, 1).Value, "0304", False)
               End If
            Case "HDA"
                temp = alterar_ME15(Cell.Value, Cell.Offset(0, 1).Value, "0212", False)
            Case "HCA"
                temp = alterar_ME15(Cell.Value, Cell.Offset(0, 1).Value, "0304", False)
            Case Else
                msgbox ("Item " & Cell.Value & " nao foi selecionada a filial para ser cancelada")
            End Select
          Select Case temp
          
               Case 0
                     Cell.Offset(0, 3).Value = "Descancelado"
               Case 1
                     Cell.Offset(0, 3).Value = "Sem no centro"
               Case 2
                     Cell.Offset(0, 3).Value = "Reginfo Nao existe"
                     End Select
          End If
        Next

        ' Sair da Transacao
        session.findById("wnd[0]").sendVKey 3

End Sub


Function cancelar_me01(material As Double, fornecedor As String, centro As String) As Integer
        
    With session

        ' Abrir o cÃ³digo de material para cancelar
        .findById("wnd[0]/usr/ctxtEORD-MATNR").Text = material
        .findById("wnd[0]/usr/ctxtEORD-WERKS").Text = centro
        .findById("wnd[0]").sendVKey 0
    
    End With
     If InStr(session.findById("wnd[0]/sbar/pane[0]").Text, "Bloqueado Somente Entrada") Then
             cancelar_me01 = 0
             Exit Function
          End If
              ' Checar se o fornecedor estÃ¡ conforme a planilha, e cancelar se estiver (Evita cancelar outros fornecedores)
    If session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,0]").Text = fornecedor Then
        With session
            .findById("wnd[0]/usr/tblSAPLMEORTC_0205").getAbsoluteRow(0).Selected = True
            .findById("wnd[0]").sendVKey 14
                .findById("wnd[1]/usr/btnSPOP-OPTION1").press
            .findById("wnd[0]").sendVKey 11
        End With
    Else
        session.findById("wnd[0]").sendVKey 3
    End If
     cancelar_me01 = 1
    End Function

Function alterar_ME15(material As Double, fornecedor As String, centro As String, cancelar As Boolean)

    With session
        .findById("wnd[0]/usr/ctxtEINA-LIFNR").Text = fornecedor
        .findById("wnd[0]/usr/ctxtEINA-MATNR").Text = material
        .findById("wnd[0]/usr/ctxtEINE-EKORG").Text = "1500"
        .findById("wnd[0]/usr/ctxtEINE-WERKS").Text = centro
        .findById("wnd[0]").sendVKey 0
     End With
     If InStr(session.findById("wnd[0]/sbar/pane[0]").Text, "dados de organiz") Then
             alterar_ME15 = 0
             session.findById("wnd[0]").sendVKey 3
             Exit Function
          End If
     If InStr(session.findById("wnd[0]/sbar/pane[0]").Text, "o existe") Then
             alterar_ME15 = 2
             Exit Function
          End If
     
    If cancelar Then
        ' Flaggar os campos de cancelamento
        session.findById("wnd[0]/usr/chkEINA-LOEKZ").Selected = True
        session.findById("wnd[0]/usr/chkEINE-LOEKZ").Selected = True
    Else
        session.findById("wnd[0]/usr/chkEINA-LOEKZ").Selected = False
        session.findById("wnd[0]/usr/chkEINE-LOEKZ").Selected = False
    End If
    session.findById("wnd[0]").sendVKey 11
     alterar_ME15 = 1
    End Function
