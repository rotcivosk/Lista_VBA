Sub PDCA_teste1()
    
    call abrir_sap

    With session
        'Abre ME5A
        .findById("wnd[0]/tbar[0]/okcd").Text = "ME5A"
        .findById("wnd[0]").sendVKey 0
        
        'Seleciona os filtros
        .findById("wnd[0]").sendVKey 17
            .findById("wnd[1]/tbar[0]/btn[8]").press
    
        'Roda o relatório
        .findById("wnd[0]").sendVKey 8
        Call exportar_clipboardSAP(False)
   
        'Sai da tela
        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
        
    End With
    
    Sheets("Planilha1").Activate
    
    Call export_Formatar_PlanilhaSAP
   
    Dim matRequisicao As Variant
    Dim matTemp As Variant
    Dim i As Long, j As Long
    Dim found As Boolean

    Sheets("Planilha1").Activate
    matRequisicao = Range("B2:W" & Range("B" & Rows.Count).End(xlUp).Row).Value

    Sheets("Planilha2").Activate
    matTemp = Range("B6:J" & Range("B" & Rows.Count).End(xlUp).Row).Value

    Sheets("Planilha3").Activate
    For i = 1 To UBound(matRequisicao, 1)
        found = False
        For j = 1 To UBound(matTemp, 1)
            If matRequisicao(i, 1) = matTemp(j, 1) And matRequisicao(i, 2) = matTemp(j, 2) Then
                Cells(i, "A").Value = matTemp(j, 3)
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            Cells(i, "A").Value = "Não encontrado"
        End If
    Next i
End Sub