Attribute VB_Name = "Importar_JDE"

Sub Importar_Catalogo_JDE()

    
    Dim fornecedor As Double
   
    Call Abrir_Chrome("http://sahdamvpjde009.sa.mds.honda.com:71/jde/E1Menu.maf?jdeLoginAction=LOGOUT&RENDER_MAFLET=E1Menu")
    Call Login_jde(user, senha)
    
    '----CATÁLOGO----
    call Abrir_tela_fav("Manutencao Catalogo de Precos")
    
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

' Importar a planilha do Catálogo

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
    call Abrir_tela_fav("Tela de Follow Pedidos Improdutivos")    


    For i = 0 To 9
        Call alterar_campo("C0_20", CStr(tip_ped(i)), "ID") ' OP, OL, OM, OS
        Call alterar_campo("C0_26", CStr(filial(i)), "ID")' 1 - 05001, 2 - 05998, 3 - 10001, 4 - 10998
        Call alterar_campo("C0_231", dt_ini, "ID") '
        Call alterar_campo("C0_233", dt_fin, "ID")
        driver.FindElementById("hc_Find").Click
        Call carregar_Exportar_JDE
        call copiar_Temp_para_Pedidos
    Next
    
    Call fechar_Chrome

    End Function
' Importar os pedidos