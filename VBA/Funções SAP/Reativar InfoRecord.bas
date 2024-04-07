Sub Descancelar_RegInfo()
'
' Cancelamento de Reginfo na ME15 e ME01
'
call Abrir_SAP

    ' Levantar a lista a ser trabalhada
    Dim lista As Range
    Set lista = Range(Range("E10"), Range("E10").End(xlDown))
    
    ' Abrir a transação ME15
    session.findById("wnd[0]/tbar[0]/okcd").Text = "me15"
    session.findById("wnd[0]").sendVKey 0
    
    For Each cell In lista
    
   
        With session
            
            ' Abrir o item para para o primeiro centro
            .findById("wnd[0]/usr/ctxtEINA-LIFNR").Text = cell.Offset(0, 1).Value
            .findById("wnd[0]/usr/ctxtEINA-MATNR").Text = cell.Value
            .findById("wnd[0]/usr/ctxtEINE-EKORG").Text = "1500"
            .findById("wnd[0]/usr/ctxtEINE-WERKS").Text = "0212"
            .findById("wnd[0]").sendVKey 0

            ' Flaggar o campo de cancelamento e salvar
            .findById("wnd[0]/usr/chkEINA-LOEKZ").Selected = False
            .findById("wnd[0]/usr/chkEINE-LOEKZ").Selected = False
            .findById("wnd[0]").sendVKey 11

            ' Abrir o segundo centro, flaggar o campo de cancelamento e salvar
            .findById("wnd[0]/usr/ctxtEINE-WERKS").Text = "0304"
            .findById("wnd[0]").sendVKey 0
            .findById("wnd[0]/usr/chkEINE-LOEKZ").Selected = False
            .findById("wnd[0]").sendVKey 11
        End With                        
        
    Next

    ' Sair da Transação
    session.findById("wnd[0]").sendVKey 3

End Sub