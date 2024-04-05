Sub Importar_Catalogo_JDE()
    
    Dim fornecedor As Double
   
    Call Abrir_Chrome("http://sahdamvpjde009.sa.mds.honda.com:71/jde/E1Menu.maf?jdeLoginAction=LOGOUT&RENDER_MAFLET=E1Menu")
    Call Login_jde(user, senha)
    
    '----CATÁLOGO----
    call Abrir_tela_fav("Manutencao Catalogo de Precos")
    Application.Wait (Now + TimeValue("0:00:08"))
    driver.SwitchToFrame 8  
    
    dim dt_ini_min as string
    dt_ini_min = ">" & dt_ini
    Call alterar_campo("qbe0_1.8", dt_ini_min, "Name") 
    call alterar_campo("C026", "DIHV*", "ID")
    driver.FindElementById("hc_Find").click

    Call carregar_Exportar_JDE
    Application.Wait (Now + TimeValue("0:00:07"))

    ThisWorkbook.Activate
    Sheets("Catálogo").Activate
    Call pull_Book1xls

    Call fechar_Chrome

End Sub
