Attribute VB_Name = "M_5_Exportar_SAP"
Sub Atualizar_SAP()

    Dim fornecedor As Double
    Dim region As Range
    
    call Abrir_SAP
    
    'EXCEL
    'Windows("Planilha de Duplo Check_2.xlsm").Activate
    Sheets("0304").Columns("B:R").ClearContents
    Sheets("0212").Columns("B:R").ClearContents
    fornecedor = Sheets("Tela Principal").Range("L5").Value
       
    With session
       
        'Abrir o ZI9
        .findById("wnd[0]/tbar[0]/okcd").Text = "zi9_mm_reginfo"
        .findById("wnd[0]").sendVKey 0
        
        'Add o 1500, o fornecedor e o Centro 0212
        .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/ctxtSEKORG").Text = "1500"
        .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/ctxtSLIFNR").Text = fornecedor
        .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/ctxtSWERKS-LOW").Text = "0212"
        
        'Roda
        .findById("wnd[0]").sendVKey 8
        
        'Exportar para a clipboard
        .findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
        .findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").selectContextMenuItem "&PC"
            .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
            .findById("wnd[1]/tbar[0]/btn[0]").press
    End With
    
    Application.Wait (Now + TimeValue("0:00:02"))
    
    'Cola no Excel
    Windows("Planilha de Duplo Check_2.xlsm").Activate
    Sheets("0212").Range("B1").PasteSpecial
    Sheets("0212").Range("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="|", TrailingMinusNumbers:=True
    
    With session
        'Sair da tela
        .findById("wnd[0]/tbar[0]/btn[3]").press
        
        'Mudar para o 0304
        .findById("wnd[0]/usr/tabsTBS_100/tabpTBS_100_FC1/ssubTBS_100_SCA:ZI9_MM_REGINFO:0101/subSBS_0104:ZI9_MM_REGINFO:0104/ctxtSWERKS-LOW").Text = "0304"
        .findById("wnd[0]/tbar[1]/btn[8]").press
        
        'Exportar o Relatório
        .findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
        .findById("wnd[0]/usr/cntlCONT_106/shellcont/shell").selectContextMenuItem "&PC"
            .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
            .findById("wnd[1]/tbar[0]/btn[0]").press
    End With
    
    Application.Wait (Now + TimeValue("0:00:02"))
    
    'Sair da tela
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    'Cola no Excel
    Windows("Planilha de Duplo Check_2.xlsm").Activate
    Sheets("0304").Range("B1").PasteSpecial
    
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    

End Sub
