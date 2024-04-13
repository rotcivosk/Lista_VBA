Attribute VB_Name = "M_SAP_Gerar carga de reajuste"
Sub Gerar_carga_reajuste()
        Dim n As Integer

    ' Selecionar a lista
        Dim lista As Range, item_ini as Range
        Set item_ini = Range("G10")
        if item_ini.Offset(1,0) = "" Then
            Set lista = item_ini
        Else
            Set lista = Range(item_ini, item_ini.End(xlDown))
        end if 

    ' Abrir SAP
        call Abrir_SAP
        
    ' Abrir a transação
    With session
        ' Abrir o ZI9
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nzi9_mm_reginfo"
        .findById("wnd[0]").sendVKey 0
        
        ' Add o 1500, o fornecedor e o Centro
        .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/ctxtSEKORG").Text = "1500"
        .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/ctxtSLIFNR").Text = Range("I6").Value
        .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/ctxtSWERKS-LOW").Text = Range("I7").Value
        End With
    
    ' Selecionar os materiais e roda
    lista.Copy
    With session
        .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/btn%_SMATNR_%_APP_%-VALU_PUSH").press
            .findById("wnd[1]/tbar[0]/btn[24]").press
            .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]").sendVKey 8
        End With

    ' Adiciona os valores 
    n = 0
    For Each cell In lista
        session.findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").modifyCell n, "ZPB0", cell.Offset(0, 1).Value
        n = n + 1
        session.findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").triggerModified
        Next
        
    ' Exportar e copiar o nome
        session.findById("wnd[0]/usr/txtCPO_CENTRO").Text = Range("I7").Value
        session.findById("wnd[0]/usr/txtCPO_TEXT").Text = Range("H8").Value
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        item_ini.Formula = mid(session.findById("wnd[0]/sbar").Text,28,4)
        session.findById("wnd[0]/tbar[0]/btn[3]").press
    End Sub
