For j = 0 To session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").RowCount - 1
    val_temp = session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").getcellvalue(j, "BITM_DESCR")
    if val_temp = "/Victor" Then
        result = J
    end if
Next




Sub Relato_SAP_PDF()
    
Call Abrir_SAP 'Chamar a sub abrir sap

'Início do comando:
Dim Req As Double
Dim Senha As String
Dim totalcoluna As Integer, Tot_Req As Integer, i As Integer, u As Integer
Dim Texto As String
Dim valor_atual As Integer

'Seleciona req base lista excel
    Tot_Req = Application.WorksheetFunction.Count(Range(Range("B3"), Range("B3").End(xlDown)))
    valor_atual = 0
    
'Abre Me53N
    session.findById("wnd[0]/tbar[0]/okcd").Text = "me53n"
    session.findById("wnd[0]").sendVKey 0
    
' Otimiza o excel
    Application.Calculation = xlCalculationManual
    UserForm2.Show vbModeless
    For Each cell In Range(Range("B3"), Range("B3").End(xlDown))
        
        If cell.Offset(0, 1) <> "" Then
            
        Else
        
            With session
               'Abre a Requisição
               .findById("wnd[0]").sendVKey 17
                   .findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN").Text = cell.Value
                   .findById("wnd[1]").sendVKey 0
            
               'Abre a pasta de Anexos
               .findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
               .findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_VIEW_ATTA"
            End With
            Texto = ""
            'Checa se Tem anexo, e se tiver, envia copia o nome
            If InStr(session.findById("wnd[0]/sbar/pane[0]").Text, "Serviço <'Lista de anexos'> indisponível") Then
                Texto = "Sem Anexo"

            Else
                'totalcoluna = session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").RowCount - 1
                For j = 0 To session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").RowCount - 1
                    If j <> 0 Then Texto = Texto & Chr(10)
                    Texto = Texto & session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").getcellvalue(j, "BITM_DESCR")
                Next j
                session.findById("wnd[1]").sendVKey 12
            End If
            cell.Offset(0, 1).FormulaR1C1 = Texto
        End If
        valor_atual = valor_atual + 1
        progress (valor_atual / Tot_Req)
        
    Next

    Application.Calculation = xlCalculationAutomatic
End Sub