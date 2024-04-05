Function Importar_Follow_JDE(dt_ini As String, dt_fin As String)
    
    Dim user As String, senha As String, i As Integer
    'Call limparPlan("Pedidos emitidos JDE")
    '---TELA DE FOLLOW----
    Call Abrir_Chrome("http://sahdamvpjde009.sa.mds.honda.com:71/jde/E1Menu.maf?jdeLoginAction=LOGOUT&RENDER_MAFLET=E1Menu")
    Call Login_jde(user, senha)
     
    Dim tip_ped() As Variant, filial() As Variant
     
    tip_ped = Array("OP", "OP", "OP", "OP", "OL", "OL", "OM", "OM", "OS", "OS")
     filial = Array("05001", "10001", "05998", "10998", "05001", "10001", "05001", "10001", "05001", "10001")

    ' Abrir tela de follow
    With driver
        .FindElementById("drop_fav_menus").Click
        .FindElementByXPath("//div[3]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td/span").Click
        .FindElementByLinkText("Tela de Follow Pedidos Improdutivos").Click
    End With
    'If esperar_carregar(" Tela de Follow Pedidos Improdutivos - Consulta RequisiÃ§Ã£o / Pedido / Fornecedor") Then Stop
    
    driver.SwitchToFrame 8

    For i = 0 To 9
        Call alterar_campo("C0_20", CStr(tip_ped(i)), "ID") ' OP, OL, OM, OS
        Call alterar_campo("C0_26", CStr(filial(i)), "ID")' 1 - 05001, 2 - 05998, 3 - 10001, 4 - 10998
        Call alterar_campo("C0_231", dt_ini, "ID") '
        Call alterar_campo("C0_233", dt_fin, "ID")
        driver.FindElementById("hc_Find").Click
        Call carregar_Exportar_JDE
    Next
    
    Call fechar_Chrome

    Sheets("Macros").Activate
End Function
