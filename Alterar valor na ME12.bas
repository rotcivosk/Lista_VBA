Sub Alterar_Valor_ME12()
'
' Alterar o Preço Unitário de um InfoRecord com base em uma lista na transação ME12
'
'
'
    ' Check para confirmação
    If MsgBox("Tem certeza que quer cadastrar? Não pode ter menos de 2 e ele vai salvar um novo", vbYesNo + vbQuestion, "Confirmation") = vbYes Then
    Else
        Range("W2").FormulaR1C1 = "No"
        Exit Sub
    End If


    call Abrir_SAP

    ' Selecionar a lista de itens a ser cancelado
    Dim lista As Range
    Set lista = Range(Range("L10"), Range("L10").End(xlDown))
    
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
