Sub Alterar_IVA_ME12()
'
' Alterar_IVA_ME12 Macro
'
'
'
call Abrir_SAP


    Dim lista As Range
    Set lista = Range(Range("H10"), Range("H10").End(xlDown))
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "me12"
    session.findById("wnd[0]").sendVKey 0
    
    For Each cell In lista
        With session
        
            'Abrir item para cancelar
            .findById("wnd[0]/usr/ctxtEINA-LIFNR").Text = cell.Offset(0, 1).Value
            .findById("wnd[0]/usr/ctxtEINA-MATNR").Text = cell.Value
            .findById("wnd[0]/usr/ctxtEINE-EKORG").Text = "1500"
            .findById("wnd[0]/usr/ctxtEINE-WERKS").Text = "0212"
            .findById("wnd[0]").sendVKey 0
            .findById("wnd[0]").sendVKey 0
            .findById("wnd[0]/usr/ctxtEINE-MWSKZ").Text = cell.Offset(0, 2).Value
            .findById("wnd[0]").sendVKey 11
            
            'Abrir o 0304
            .findById("wnd[0]/usr/ctxtEINE-WERKS").Text = "0304"
            .findById("wnd[0]").sendVKey 0
            .findById("wnd[0]").sendVKey 0
            .findById("wnd[0]/usr/ctxtEINE-MWSKZ").Text = cell.Offset(0, 2).Value
            .findById("wnd[0]").sendVKey 11
        End With
    Next
    session.findById("wnd[0]/tbar[0]/btn[3]").press

End Sub
