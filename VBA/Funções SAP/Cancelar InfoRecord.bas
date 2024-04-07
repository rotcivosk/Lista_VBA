Sub Cancelar_InfoRecord()
'
' Cancelamento de Inforecord nas transações ME15 e ME01
'
call Abrir_SAP

    ' Selecionar a lista a ser cancelada
    Dim lista As Range
    Set lista = Range(Range("B10"), Range("B10").End(xlDown))

    ' Abrir Transação ME15
    session.findById("wnd[0]/tbar[0]/okcd").Text = "me15"
    session.findById("wnd[0]").sendVKey 0


    For Each cell In lista
     
      With session
          ' Abrir o código de material para cancelar
          .findById("wnd[0]/usr/ctxtEINA-LIFNR").Text = cell.Offset(0, 1).Value
          .findById("wnd[0]/usr/ctxtEINA-MATNR").Text = cell.Value
          .findById("wnd[0]/usr/ctxtEINE-EKORG").Text = "1500"
          .findById("wnd[0]/usr/ctxtEINE-WERKS").Text = "0212"
          .findById("wnd[0]").sendVKey 0

          ' Flaggar os campos de cancelamento
          .findById("wnd[0]/usr/chkEINA-LOEKZ").Selected = True
          .findById("wnd[0]/usr/chkEINE-LOEKZ").Selected = True
          .findById("wnd[0]").sendVKey 11

          ' Alterar para o novo centro e flaggar os campos de cancelamento
          .findById("wnd[0]/usr/ctxtEINE-WERKS").Text = "0304"
          .findById("wnd[0]").sendVKey 0
          .findById("wnd[0]/usr/chkEINE-LOEKZ").Selected = True
          .findById("wnd[0]").sendVKey 11
      End With
    Next

    ' Abrir transação ME01
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme01"
    session.findById("wnd[0]").sendVKey 0


    For Each cell In lista
    
        fornecedor = cell.Offset(0, 1).Value
        With session
            ' Abrir o código de material para cancelar
            .findById("wnd[0]/usr/ctxtEORD-MATNR").Text = cell.Value
            .findById("wnd[0]/usr/ctxtEORD-WERKS").Text = "0212"
            .findById("wnd[0]").sendVKey 0
        End With

        ' Checar se o fornecedor está conforme a planilha, e cancelar se estiver (Evita cancelar outros fornecedores)
        If session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,0]").Text = cell.Offset(0, 1).Value Then
            With session
                .findById("wnd[0]/usr/tblSAPLMEORTC_0205").getAbsoluteRow(0).Selected = True
                .findById("wnd[0]").sendVKey 14
                    .findById("wnd[1]/usr/btnSPOP-OPTION1").press
                .findById("wnd[0]").sendVKey 11
            End With
        Else
            session.findById("wnd[0]").sendVKey 3
        End If

        ' Alterar para o novo centro e flaggar os campos de cancelamento
        session.findById("wnd[0]/usr/ctxtEORD-WERKS").Text = "0304"
        session.findById("wnd[0]").sendVKey 0


        ' Checar se o fornecedor está conforme a planilha, e cancelar se estiver (Evita cancelar outros fornecedores)
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
    Next
'
    session.findById("wnd[0]").sendVKey 3

End Sub
