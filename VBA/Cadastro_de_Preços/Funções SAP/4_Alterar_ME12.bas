Attribute VB_Name = "M_4_Alterar_ME12"
Sub Alterar_IVA_ME12()
    '
    call Abrir_SAP

    Dim lista As Range, item_ini as Range
        Set item_ini = Range("R10")
        if item_ini.Offset(1,0) = "" Then
            Set lista = item_ini
        Else
            Set lista = Range(item_ini, item_ini.End(xlDown))
        end if 
    ' Selecionar a lista
    
       
    session.findById("wnd[0]/tbar[0]/okcd").Text = "me12"
    session.findById("wnd[0]").sendVKey 0
    
    For Each cell In lista
        With session
            ' Alterar o item
            .findById("wnd[0]/usr/ctxtEINA-LIFNR").Text = cell.Offset(0, 1).Value
            .findById("wnd[0]/usr/ctxtEINA-MATNR").Text = cell.Value
            .findById("wnd[0]/usr/ctxtEINE-EKORG").Text = "1500"
            .findById("wnd[0]/usr/ctxtEINE-WERKS").Text = "0212"
            .findById("wnd[0]").sendVKey 0
            .findById("wnd[0]").sendVKey 0
            .findById("wnd[0]/usr/ctxtEINE-MWSKZ").Text = cell.Offset(0, 2).Value
            .findById("wnd[0]").sendVKey 11
            
            ' Abrir o 0304
            .findById("wnd[0]/usr/ctxtEINE-WERKS").Text = "0304"
            .findById("wnd[0]").sendVKey 0
            .findById("wnd[0]").sendVKey 0
            .findById("wnd[0]/usr/ctxtEINE-MWSKZ").Text = cell.Offset(0, 2).Value
            .findById("wnd[0]").sendVKey 11
            End With
        Next
    session.findById("wnd[0]/tbar[0]/btn[3]").press

    End Sub

Sub Alterar_Valor_ME12()
'
' Alterar o Preço Unitário de um InfoRecord com base em uma lista na transação ME12
'
'
'
    ' Check para confirmação
    If MsgBox("Confirmar que gostaria de alterar o valor do IVA abaixo", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub

    call Abrir_SAP

    Dim lista As Range, item_ini as Range
        Set item_ini = Range("V10")
        if item_ini.Offset(1,0) = "" Then
            Set lista = item_ini
        Else
            Set lista = Range(item_ini, item_ini.End(xlDown))
        end if 
    ' Selecionar a lista
    
    ' Abrir a transação ME12
    session.findById("wnd[0]/tbar[0]/okcd").Text = "me12"
    session.findById("wnd[0]").sendVKey 0
    
    For Each cell In lista
        With session
        
            ' Abrir o InfoRecord
            .findById("wnd[0]/usr/ctxtEINA-LIFNR").Text = cell.Offset(0, 1).Value
            .findById("wnd[0]/usr/ctxtEINA-MATNR").Text = cell.Value
            .findById("wnd[0]/usr/ctxtEINE-EKORG").Text = "1500"
            .findById("wnd[0]/usr/ctxtEINE-WERKS").Text = cell.Offset(0, 3).Value
            .findById("wnd[0]").sendVKey 0
            
            ' Alterar o valor (Criando um novo)
            .findById("wnd[0]").sendVKey 8
                .findById("wnd[1]").sendVKey 7
            .findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/txtKONP-KBETR[2,0]").Text = cell.Offset(0, 2).Value
            .findById("wnd[0]").sendVKey 11    
            End With
        Next

    ' Sair da transação
    session.findById("wnd[0]/tbar[0]/btn[3]").press

    End Sub
