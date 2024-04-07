
Sub Relato_SAP_PDF()
    
Call Abrir_SAP 'Chamar a sub abrir sap

'Início do comando:
Dim totalcoluna As Integer, Tot_Req As Integer, i As Integer, u As Integer
Dim Texto As String

'Seleciona req base lista excel
    Tot_Req = Application.WorksheetFunction.Count(Range(Range("B3"), Range("B3").End(xlDown)))

'Abre Me53N
    session.findById("wnd[0]/tbar[0]/okcd").Text = "me53n"
    session.findById("wnd[0]").sendVKey 0
    
' Otimiza o excel
    Application.Calculation = xlCalculationManual
    
    i = 0
    For Each Cell In Range(Range("B3"), Range("B3").End(xlDown))
        i = i + 1
        If Cell.Offset(0, 1) <> "" Then
           
        Else
        
            With session
               'Abre a Requisição
               .findById("wnd[0]").sendVKey 17
                   .findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN").Text = Cell.Value
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
                For j = 1 To session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").RowCount
                    If j <> 1 Then Texto = Texto & Chr(10)
                    Texto = Texto & session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").getcellvalue(j - 1, "BITM_DESCR")
                Next j
                session.findById("wnd[1]").sendVKey 12
            End If
            Cell.Offset(0, 1).FormulaR1C1 = Texto
            Call barra_status(i, Tot_Req)
        End If
    Next

    Application.StatusBar = False

    Application.Calculation = xlCalculationAutomatic
End Sub
