Public user As String, senha As String
Sub testando()

     user = "Sb048948"
     senha = "Compras@98"

     Call Importar_Catalogo_JDE("05/04/2024")
     Call Importar_Follow_JDE("27/03/2024", "05/04/2024")


End Sub

Sub Importar_Catalogo_JDE(data_input As String)
   
    Call Abrir_Chrome("http://sahdamvpjde009.sa.mds.honda.com:71/jde/E1Menu.maf?jdeLoginAction=LOGOUT&RENDER_MAFLET=E1Menu")
    Call Login_jde(user, senha)
    Call Abrir_tela_fav("Manutencao Catalogo de Precos")
    
    Dim dt_ini_min As String
    dt_ini_min = " > " & data_input
    
    Call alterar_campo("qbe0_1.8", dt_ini_min, "Name")
    Call alterar_campo("C0_26", "DIVH*", "ID")
    driver.FindElementById("hc_Find").Click
    Call wait_loading_page

    Call carregar_Exportar_JDE
    Application.Wait (Now + TimeValue("0:00:07"))

    ThisWorkbook.Activate
    Sheets("Catalogo").Activate
    'Call pull_Book1xls

    Call fechar_Chrome

    End Sub
' Importar a planilha do CatÃ¡logo

Function Importar_Follow_JDE(dt_ini As String, dt_fin As String)
    
    Call Abrir_Chrome("http://sahdamvpjde009.sa.mds.honda.com:71/jde/E1Menu.maf?jdeLoginAction=LOGOUT&RENDER_MAFLET=E1Menu")
    Call Login_jde(user, senha)
     
    Dim tip_ped() As Variant, filial() As Variant
     
    tip_ped = Array("OP", "OP", "OP", "OP", "OL", "OL", "OM", "OM", "OS", "OS")
     filial = Array("05001", "10001", "05998", "10998", "05001", "10001", "05001", "10001", "05001", "10001")

    ' Abrir tela de follow
    Call Abrir_tela_fav("Tela de Follow Pedidos Improdutivos")


    For i = 0 To 9
        Call alterar_campo("C0_20", CStr(tip_ped(i)), "ID") ' OP, OL, OM, OS
        Call alterar_campo("C0_26", CStr(filial(i)), "ID") ' 1 - 05001, 2 - 05998, 3 - 10001, 4 - 10998
        Call alterar_campo("C0_231", dt_ini, "ID") '
        Call alterar_campo("C0_233", dt_fin, "ID")
        driver.FindElementById("hc_Find").Click
        Call carregar_Exportar_JDE
        Call copiar_Temp_para_Pedidos
    Next
    
    Call fechar_Chrome

    End Function
' Importar os pedidos
