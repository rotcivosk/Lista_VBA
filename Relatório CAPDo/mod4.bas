'Relatório base para Requisições em aberto
'Special thanks to whoever made the coffee when i did this
'Índice:
'Rel_SAP_meuAcesso
' ----BASE----
' ----ME5A----
' ----ME2N----
'Rel_SAP_semacesso
' ----EORD----
' ----EINA----
' ----EINE----


Sub Rel_SAP_meuAcesso()




'----BASE----


    Call Abrir_SAP
    
    Dim dt_ini, dt_fin As String

'--RELATÓRIOS--
    
    dt_ini = Range("C3").Value
    dt_fin = Range("C4").Value
    
    ' ----ME5A----
  
    With session
        'Abre ME5A
        .findById("wnd[0]/tbar[0]/okcd").Text = "me5a"
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
    
    Sheets("ME5A").Activate
    
    Call export_Formatar_PlanilhaSAP
    Columns("F:F").NumberFormat = "General"


    ' ---- ME2N ----
    
    
    
    With session

        'Abre a ME2N
        .findById("wnd[0]/tbar[0]/okcd").Text = "me2n"
        .findById("wnd[0]").sendVKey 0
        
        'Seleciona os filtrons
        .findById("wnd[0]/usr/ctxtEN_EBELN-LOW").Text = "45*"
        .findById("wnd[0]/usr/ctxtLISTU").Text = "alv"
        'Datas
        .findById("wnd[0]/usr/ctxtS_BEDAT-LOW").Text = dt_ini
        .findById("wnd[0]/usr/ctxtS_BEDAT-HIGH").Text = dt_fin
        .findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").press
            .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "0212"
            .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "0304"
            .findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "0232"
            .findById("wnd[1]").sendVKey 8

        'Rodar
        .findById("wnd[0]").sendVKey 8
        .findById("wnd[0]/tbar[1]/btn[23]").press
        
        Call exportar_clipboardSAP(False)

        .findById("wnd[0]/tbar[0]/btn[3]").press
        .findById("wnd[0]/tbar[0]/btn[3]").press
    
    End With

    Sheets("ME2N").Activate
    Call export_Formatar_PlanilhaSAP
    Sheets("Macros").Activate
End Sub



Sub Rel_SAP_semacesso()


    Call Abrir_SAP
    
    ' ----EORD----

    Sheets("ME5A").Activate

    Range(Range("F2"), Range("F2").End(xlDown)).Copy
    
    With session
    
        'Open SE16N - EORD
        .findById("wnd[0]/tbar[0]/okcd").Text = "se16n"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/usr/ctxtGD-TAB").Text = "eord"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/usr/txtGD-MAX_LINES").Text = ""
        'Aplicar filtros
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").SetFocus
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press
            .findById("wnd[1]/tbar[0]/btn[24]").press
            .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]/tbar[1]/btn[8]").press
        'Exportar
        
        Call exportar_clipboardSAP(True)
        .findById("wnd[0]/tbar[0]/btn[3]").press
    End With
    Application.Wait (Now + TimeValue("0:00:01"))
    Sheets("EORD").Activate
    Call export_Formatar_PlanilhaSAP
    
    
    ' ----EINA----
    
    Range(Range("A2"), Range("a2").End(xlDown)).Copy
    With session
        'Cabeçalho
        .findById("wnd[0]/usr/ctxtGD-TAB").Text = "eina"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/usr/txtGD-MAX_LINES").Text = ""
        'Material
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,2]").SetFocus
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,2]").press
            .findById("wnd[1]/tbar[0]/btn[24]").press
            .findById("wnd[1]/tbar[0]/btn[8]").press
    End With

    'Fornecedor
    Range(Range("H2"), Range("H2").End(xlDown)).Copy
    With session
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,4]").SetFocus
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,4]").press
            .findById("wnd[1]/tbar[0]/btn[24]").press
            .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]/tbar[1]/btn[8]").press
        'Export
        Call exportar_clipboardSAP(True)
        .findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").setCurrentCell 11, "UMREZ"
        .findById("wnd[0]").sendVKey 3
    End With
    Application.Wait (Now + TimeValue("0:00:01"))
    Sheets("EINA").Activate
    Call export_Formatar_PlanilhaSAP
    
    
    
    
    '----EINE----
    
    Range(Range("A2"), Range("A2").End(xlDown)).Copy
    With session
        .findById("wnd[0]/usr/ctxtGD-TAB").Text = "eine"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").SetFocus
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press
            .findById("wnd[1]/tbar[0]/btn[24]").press
            .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,4]").SetFocus
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,4]").press
            .findById("wnd[1]/usr/tblSAPLSE16NMULTI_TC/ctxtGS_MULTI_SELECT-LOW[1,0]").Text = "0212"
            .findById("wnd[1]/usr/tblSAPLSE16NMULTI_TC/ctxtGS_MULTI_SELECT-LOW[1,1]").Text = "0304"
            .findById("wnd[1]/usr/tblSAPLSE16NMULTI_TC/ctxtGS_MULTI_SELECT-LOW[1,2]").Text = "0232"
            .findById("wnd[1]").sendVKey 8
        .findById("wnd[0]").sendVKey 8

        Call exportar_clipboardSAP(True)
        Application.Wait (Now + TimeValue("0:00:01"))
    End With
    Sheets("EINE").Activate
    Call export_Formatar_PlanilhaSAP




    
    '---CDPOS---


    'Puxar as REQs
        
    Sheets("Temp").Activate
    Columns("B:B").ClearContents
    
    Sheets("ME5A").Activate
    Range(Range("B2"), Range("B2").End(xlDown)).Copy
    
    Sheets("Temp").Activate
    Range("B2").Select
    ActiveSheet.Paste
    
    Sheets("ME2N").Activate
    Range(Range("C2"), Range("C2").End(xlDown)).Copy

    Sheets("Temp").Activate
    Range("B2").Select
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    ActiveSheet.Paste
    
    With Range(Range("B2"), Range("B2").End(xlDown))
        .Select
        .NumberFormat = "0000000000"
        .RemoveDuplicates Columns:=1, Header:=xlNo
    End With
    ActiveSheet.Calculate
    Application.Wait (Now + TimeValue("0:00:01"))
    Range(Range("B2"), Range("B2").End(xlDown)).Copy
    
    With session
            
        .findById("wnd[0]").sendVKey 3
        .findById("wnd[0]/usr/ctxtGD-TAB").Text = "cdpos"
        .findById("wnd[0]").sendVKey 0
        'Rodar
        .findById("wnd[0]").sendVKey 6
            .findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").Text = "APROV_REQ"
            .findById("wnd[1]/usr/txtGS_SE16N_LT-UNAME").Text = "SB048948"
            .findById("wnd[1]").sendVKey 0
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,2]").SetFocus
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,2]").press
            .findById("wnd[1]").sendVKey 24
            .findById("wnd[1]").sendVKey 8
            
        .findById("wnd[0]/tbar[1]/btn[8]").press
        
        Call exportar_clipboardSAP(True)

    End With
    Application.Wait (Now + TimeValue("0:00:03"))
    Sheets("CDPOS").Activate
    Call export_Formatar_PlanilhaSAP
    
    
    ' ----CDHDR----
    
    Range(Range("C2"), Range("C2").End(xlDown)).Copy
    With session
            
        .findById("wnd[0]").sendVKey 3
        .findById("wnd[0]/usr/ctxtGD-TAB").Text = "CDHDR"
        .findById("wnd[0]").sendVKey 0
        'Rodar
        .findById("wnd[0]").sendVKey 6
            .findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").Text = "APROV_REQ"
            .findById("wnd[1]/usr/txtGS_SE16N_LT-UNAME").Text = "SB048948"
            .findById("wnd[1]").sendVKey 0
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,3]").SetFocus
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,3]").press
            .findById("wnd[1]").sendVKey 24
            .findById("wnd[1]").sendVKey 8
        .findById("wnd[0]").sendVKey 8
        
        Call exportar_clipboardSAP(True)

    End With
    Application.Wait (Now + TimeValue("0:00:01"))
    Sheets("CDHDR").Activate
    Call export_Formatar_PlanilhaSAP



       
    ' ----EKKO----
   
    Sheets("ME2N").Activate
    Columns("A:A").NumberFormat = "General"
    ActiveSheet.Calculate
    Range(Range("A2"), Range("A2").End(xlDown)).Copy
   
   With session
            
        .findById("wnd[0]").sendVKey 3
        .findById("wnd[0]/usr/ctxtGD-TAB").Text = "EKKO"
        .findById("wnd[0]").sendVKey 0
        'Rodar
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").SetFocus
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press
            .findById("wnd[1]").sendVKey 24
            .findById("wnd[1]").sendVKey 8
        .findById("wnd[0]").sendVKey 8
        
        Call exportar_clipboardSAP(True)

    End With
    Application.Wait (Now + TimeValue("0:00:03"))
    Sheets("EKKO").Activate
    Call export_Formatar_PlanilhaSAP
   
   
   
    ' ----KONV----

    Columns("C:C").NumberFormat = "General"
    ActiveSheet.Calculate
    Range(Range("C2"), Range("C2").End(xlDown)).Copy
   
   With session

        .findById("wnd[0]").sendVKey 3
        .findById("wnd[0]/usr/ctxtGD-TAB").Text = "KONV"
        .findById("wnd[0]").sendVKey 0
        'Rodar
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").SetFocus
        .findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press
            .findById("wnd[1]").sendVKey 24
            .findById("wnd[1]").sendVKey 8
        .findById("wnd[0]").sendVKey 8
        
        Call exportar_clipboardSAP(True)

    End With
    Application.Wait (Now + TimeValue("0:00:03"))
    Sheets("KONV").Activate
    Call export_Formatar_PlanilhaSAP
   
   
    Sheets("Macros").Activate
End Sub
