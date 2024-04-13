Attribute VB_Name = "M_JDE_Funcoes_Gerais"


Public driver
' Lembre-se de add o Selenium Type Library em Ferramentas -> Referencias

Sub Abrir_Chrome(site As String)
    Set driver = New ChromeDriver
    Debug.Print "Abrindo PÃ¡gina "; site
    driver.Get site
    driver.Window.maximize
    End Sub
' Abrir site no Google Chrome

Sub Login_jde(user As String, senha As String)
    Debug.Print "Iniciando o Login"
    Application.Wait (Now + TimeValue("0:00:02"))
    Call alterar_campo("User", user, "ID")
    Call alterar_campo("Password", senha, "ID")
    Debug.Print "Credenciais Adicionadas"
    driver.FindElementByCss(".buttonstylenormal").Click
    Application.Wait (Now + TimeValue("0:00:03"))
    Debug.Print "Login Realizado"
    End Sub
' Fazer Login no JDE

Sub fechar_Chrome()
    driver.Quit
    Set driver = Nothing
    Debug.Print "Fechar Chrome"
    End Sub
' Fechar o chrome e encerrar o uso do driver

Function alterar_campo(campo As String, dado_inputado As String, tipo_proc_element As String)
    Select Case tipo_proc_element
    Case "ID"
        With driver
            .FindElementById(campo, timeout:=10000).Clear
            .ExecuteScript ("document.getElementById('" & campo & "').value = '" & dado_inputado & "'")
            .FindElementById(campo, timeout:=10000).SendKeys Enter
        End With
    Case "Name"
        With driver
            .FindElementByName(campo, timeout:=10000).Clear
            .ExecuteScript ("document.getElementsByName('" & campo & "')[0].value = '" & dado_inputado & "';" & "var evt = document.createEvent('HTMLEvents');" & "evt.initEvent('change', false, true);" & "document.getElementsByName('" & campo & "')[0].dispatchEvent(evt);")
            .ExecuteScript ("document.getElementsByName('" & campo & "').value = '" & dado_inputado & "'")
            .FindElementByName(campo, timeout:=10000).SendKeys Enter
        End With
    Case Else
        Debug.Print "Mandou algo de errado aÃ­ no" & campo
    End Select
    Debug.Print "Added " & dado_inputado & " no campo " & campo
    End Function
' Alterar o valor de um campo dado determinado ID ou Nome

Function esperar_carregar(texto As String) As Boolean
    n = 0
    Do
        titulo_pag = driver.FindElementById("jdeFormTitle0", timeout:=10000).Text
        n = n + 1
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop Until titulo_pag = texto Or n > 10

    esperar_carregar = False
    If n > 9 Then
        esperar_carregar = True
    End If
    End Function
' Esperar carregar uma tela em atÃ© 10 segundos

Function Abrir_tela_fav(texto As String)
    With driver
            .FindElementById("drop_fav_menus").Click
            .FindElementByXPath("//div[3]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td/span").Click
            .FindElementByLinkText(texto).Click
    End With
     Call wait_loading_page
    driver.SwitchToFrame 8
    End Function
' Selecionar uma tela do menu favoritos e abrir. NecessÃ¡rio melhor maneira de esperar do que Aplication Wait

Sub wait_loading_page()
    Dim temp, n As Integer
     n = 0
     Do While n < 15
         Debug.Print n; " Seg."
         On Error Resume Next
         temp = driver.FindElementById("ariaLog").Text
         Debug.Print temp
         On Error GoTo 0
         If InStr(temp, "carregamento de ") Then
               Debug.Print "OK"
               Exit Do
         End If
         temp = driver.FindElementByTag("body").Attribute("style")
         Debug.Print temp
         If InStr(temp, "cursor: auto") > 0 Then
               Debug.Print "OK"
               Exit Do
         End If
         Application.Wait (Now + TimeValue("0:00:01"))
         n = n + 1
     Loop
     If n > 15 Then
          MsgBox ("Timeout na Abertura de Tela")
          Stop
     End If
    End Sub
' Esperar o loading

Sub carregar_Exportar_JDE()
    Dim qtd_err As Integer
    qtd_err = 0
    
     
     On Error Resume Next
     driver.FindElementById("jdehtmlGridLast0_1").Click
     driver.FindElementById("GOTOLAST0_1").Click
     
     On Error GoTo 0
     Call wait_loading_page
     
     With driver
          .FindElementById("jdehtmlExportData0_1").Click
          .FindElementById("hc1").Click
     End With

    ThisWorkbook.Activate
    Sheets("Temp").Activate

    Application.DisplayAlerts = False
    Debug.Print "tela Limpa"
    Range(Range("A1"), ActiveCell.SpecialCells(xlLastCell)).ClearContents
    Application.Wait (Now + TimeValue("0:00:15"))
    Debug.Print "Book1.xlsx abrindo"
    Workbooks.Open ("D:\Users\sb048948\Downloads\Book1.xlsx")
    Debug.Print "Book1.xlsx aberto"
     
    Range(Range("A1"), ActiveCell.SpecialCells(xlLastCell)).Copy

    ThisWorkbook.Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Workbooks("Book1.xlsx").Close SaveChanges:=False
    Debug.Print "Book1.xlsx Fechando"
    Kill ("D:\Users\sb048948\Downloads\Book1.xlsx")
    Application.DisplayAlerts = True
     Debug.Print "Book1.xlsx Apagando"
    End Sub
' Esperar para exportar. Este cÃ³digo funciona, mas estÃ¡ longe de ser estÃ¡vel

Sub copiar_Temp_para_Pedidos()
    Sheets("Temp").Activate
    Range(Range("A2"), ActiveCell.SpecialCells(xlLastCell)).Copy
    Sheets("Pedidos Emitidos JDE").Activate
    Range("E2").End(xlDown).Select
    Selection.Offset(1, -4).Select
    ActiveSheet.Paste
    End Sub
' Copiar da planilha temp, apÃ³s o Book1.xlsx e colar na planilha pedidos emitidos JDE

