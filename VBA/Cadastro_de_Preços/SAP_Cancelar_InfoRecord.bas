Attribute VB_Name = "M_SAP_Cancelar_InfoRecord"
Sub Cancelar_InfoRecord()
    '
    ' Cancelamento de Inforecord nas transações ME15 e ME01
    '
    call Abrir_SAP

    ' Selecionar a lista a ser cancelada
        Dim lista As Range, item_ini as Range
        Set item_ini = Range("B10")
        if item_ini.Offset(1,0) = "" Then
            Set lista = item_ini
        Else
            Set lista = Range(item_ini, item_ini.End(xlDown))
        end if 

    ' Abrir Transação ME15
    session.findById("wnd[0]/tbar[0]/okcd").Text = "me15"
    session.findById("wnd[0]").sendVKey 0

    For Each cell In lista
        select case cell.Offset(0,2)
            case "AMBOS"            
                call alterar_ME15(cell.Value, cell.Offset(0,1).Value, "0212", True)
                call alterar_ME15(cell.Value, cell.Offset(0,1).Value, "0304", True)
            case "HDA"
                call alterar_ME15(cell.Value, cell.Offset(0,1).Value, "0212", True)
            case "HCA"
                call alterar_ME15(cell.Value, cell.Offset(0,1).Value, "0304", True)
            case Else
                msgbox ("Item " & cell.Value & " não foi selecionada a filial para ser cancelada")
            end select
            cell.Offset(0,3).text = "Cancelado ME15"
        Next

    ' Abrir transação ME01
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme01"
    session.findById("wnd[0]").sendVKey 0

    For Each cell In lista
        select case cell.Offset(0,2)
            case "AMBOS"            
                call cancelar_me01(cell.Value, cell.Offset(0,1).Value, "0212")
                call cancelar_me01(cell.Value, cell.Offset(0,1).Value, "0304")
            case "HDA"
                call cancelar_me01(cell.Value, cell.Offset(0,1).Value, "0212")
            case "HCA"
                call cancelar_me01(cell.Value, cell.Offset(0,1).Value, "0304")
            case Else
                msgbox ("Item " & cell.Value & " não foi selecionada a filial para ser cancelada")
            end select
            cell.Offset(0,3).text = "Cancelado"

        Next

    session.findById("wnd[0]").sendVKey 3
    End Sub

Sub Descancelar_RegInfo()
    '
    ' Cancelamento de Inforecord nas transações ME15 e ME01
    '
    call Abrir_SAP

    ' Selecionar a lista a ser cancelada
        Dim lista As Range, item_ini as Range
        Set item_ini = Range("B10")
        if item_ini.Offset(1,0) = "" Then
            Set lista = item_ini
        Else
            Set lista = Range(item_ini, item_ini.End(xlDown))
        end if 

        ' Abrir a transação ME15
        session.findById("wnd[0]/tbar[0]/okcd").Text = "me15"
        session.findById("wnd[0]").sendVKey 0
        
    For Each cell In lista
        select case cell.Offset(0,2)
            case "AMBOS"            
                call alterar_ME15(cell.Value, cell.Offset(0,1).Value, "0212", True)
                call alterar_ME15(cell.Value, cell.Offset(0,1).Value, "0304", True)
            case "HDA"
                call alterar_ME15(cell.Value, cell.Offset(0,1).Value, "0212", True)
            case "HCA"
                call alterar_ME15(cell.Value, cell.Offset(0,1).Value, "0304", True)
            case Else
                msgbox ("Item " & cell.Value & " não foi selecionada a filial para ser cancelada")
            end select
            cell.Offset(0,3).text = "Descancelado"
        Next

        ' Sair da Transação
        session.findById("wnd[0]").sendVKey 3

End Sub                                                     


function cancelar_me01(material as double, fornecedor as double, centro as text)
        
    With session

        ' Abrir o código de material para cancelar
        .findById("wnd[0]/usr/ctxtEORD-MATNR").Text = material
        .findById("wnd[0]/usr/ctxtEORD-WERKS").Text = centro
        .findById("wnd[0]").sendVKey 0
    
    End With
    
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
    end function

function alterar_ME15(material as double, fornecedor as double, centro as text, cancelar as boolean)

    with session
        .findById("wnd[0]/usr/ctxtEINA-LIFNR").Text = fornecedor
        .findById("wnd[0]/usr/ctxtEINA-MATNR").Text = material
        .findById("wnd[0]/usr/ctxtEINE-EKORG").Text = "1500"
        .findById("wnd[0]/usr/ctxtEINE-WERKS").Text = centro
        .findById("wnd[0]").sendVKey 0
    end with

    if cancelar Then
        ' Flaggar os campos de cancelamento
        session.findById("wnd[0]/usr/chkEINA-LOEKZ").Selected = True
        session.findById("wnd[0]/usr/chkEINE-LOEKZ").Selected = True
    Else
        session.findById("wnd[0]/usr/chkEINA-LOEKZ").Selected = false
        session.findById("wnd[0]/usr/chkEINE-LOEKZ").Selected = false
    end if  
    session.findById("wnd[0]").sendVKey 11

    end function