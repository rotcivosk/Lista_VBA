Attribute VB_Name = "M_Importar_OB"
Sub Importar_OB()
    
    dim user as string, senha as string
    Dim fornecedor As Double
    fornecedor = Sheets("Tela Principal").Range("L4").Value
   
    Call Abrir_Chrome("http://sahdamvpjde009.sa.mds.honda.com:71/jde/E1Menu.maf?jdeLoginAction=LOGOUT&RENDER_MAFLET=E1Menu")
    Call Login_jde(user, senha)
    Call Abrir_tela_fav("Consulta Planejamento Compras")

    ' Alterar os campos
    Call alterar_campo("C0_52", fornecedor, "ID")
    With driver
        .FindElementById("C0_62").Click
        .FindElementById("hc_Find").Click
    End With
    Call wait_loading_page

    Call carregar_Exportar_JDE
    Application.Wait (Now + TimeValue("0:00:07"))
    
    Call fechar_Chrome
    
    ThisWorkbook.Activate
        Sheets("OB").Activate
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