Sub Gerar_carga_reajuste()
'
' Macro1 Macro
'

'
    Dim n As Integer

    ' Ordenar aos Valores
    Range(Range("Q10"), Range("Q10").End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select

    call Abrir_SAP
    
    With session
           
        ' Abrir o ZI9
        .findById("wnd[0]/tbar[0]/okcd").Text = "zi9_mm_reginfo"
        .findById("wnd[0]").sendVKey 0
        
        ' Add o 1500, o fornecedor e o Centro 0212
        .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/ctxtSEKORG").Text = "1500"
        .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/ctxtSLIFNR").Text = Range("C2").Value
        .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/ctxtSWERKS-LOW").Text = Range("C3").Value
        
    End With
    
    ' Selecionar os materiais e roda
    Range(Range("Q10"), Range("Q10").End(xlDown)).Copy
    With session
        .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/btn%_SMATNR_%_APP_%-VALU_PUSH").press
            .findById("wnd[1]/tbar[0]/btn[24]").press
            .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]").sendVKey 8
    End With
            
    n = 0
    For Each cell In Range(Range("Q10"), Range("Q10").End(xlDown))
    
        session.findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").modifyCell n, "ZPB0", cell.Offset(0, 1).Value
        n = n + 1
        session.findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").triggerModified
    Next
        
    ' Exportar e copiar o nome
    session.findById("wnd[0]/usr/txtCPO_CENTRO").Text = Range("C3").Value
    session.findById("wnd[0]/usr/txtCPO_TEXT").Text = Range("F2").Value
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    Range("F4").Formula = session.findById("wnd[0]/sbar").Text
    session.findById("wnd[0]/tbar[0]/btn[3]").press
End Sub