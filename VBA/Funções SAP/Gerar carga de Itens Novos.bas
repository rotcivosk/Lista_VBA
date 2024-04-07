Sub Gerar_carga_itens_novos()
'
' Macro1 Macro
'

'
    Dim n As Integer

    'Ordenar aos Valores
    'Range(Range("X10"), Range("X10").End(xlToRight)).Select
 
 
    Abrir_SAP
    
    
    'Abrir o ZI9
    session.findById("wnd[0]/tbar[0]/okcd").Text = "zi9_mm_reginfo"
    session.findById("wnd[0]").sendVKey 0
    
    
    For Each cell In Range(Range("X10"), Range("X10").End(xlDown))
        With session
            'Add o 1500, o fornecedor e o Centro 0212
            .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC2").Select
            .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC2/ssubTBS_100_SCA:ZI9_MM_REGINFO:0102/subSBS_0105:ZI9_MM_REGINFO:0105/ctxtS_EBELN-LOW").Text = cell.Value
            .findById("wnd[0]").sendVKey 8
            
            'Rodar o relatório
            .findById("wnd[0]/usr/txtCPO_TEXT").Text = Range("F2").Value
            .findById("wnd[0]/tbar[1]/btn[8]").press
            
        End With
        cell.Offset(0, 1).Formula = Right$(Left$(session.findById("wnd[0]/sbar").Text, 31), 4)
    Next
    
    session.findById("wnd[0]/tbar[0]/btn[3]").press
End Sub