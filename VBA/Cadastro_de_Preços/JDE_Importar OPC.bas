Attribute VB_Name = "M_Importar_OPC"
Sub Importar_OPC()
    
    Dim fornecedor As Double
    fornecedor = Range("L4").Value

    dim user as string, senha as string
        Call Abrir_Chrome("http://sahdamvpjde009.sa.mds.honda.com:71/jde/E1Menu.maf?jdeLoginAction=LOGOUT&RENDER_MAFLET=E1Menu")
        Call Login_jde(user, senha)
        Call Abrir_tela_fav("Consulta Planejamento Compras")
    ' Abrir o JDE
    
    Call alterar_campo("C0_26", "DIVH*", "ID")
    Call alterar_campo("qbe0_1.8", range("T2").Value, "Name")
    driver.FindElementById("hc_Find").Click

    Call carregar_Exportar_JDE
    Application.Wait (Now + TimeValue("0:00:07"))
    Call fechar_Chrome
    
    ThisWorkbook.Activate
    Sheets("OPC").Activate
    Range("A3").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).ClearContents
    

    Call pull_Book1xls

    ThisWorkbook.Activate
    Range("A2").Copy
    Range("B1").Select
    Selection.End(xlDown).Select
    Selection.Offset(0, -1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    
End Sub