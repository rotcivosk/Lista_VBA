Attribute VB_Name = "Funcoes_JDE"

Public driver
' Lembre-se de add o Selenium Type Library em Ferramentas -> Referencias

Sub Abrir_Chrome(site As String)
    Set driver = New ChromeDriver
    driver.Get site
    driver.Window.maximize
    End Sub
' Abrir site no Google Chrome

Sub Login_jde(user As String, Senha As String)
    Application.Wait (Now + TimeValue("0:00:02"))
    call alterar_campo("User", user, "ID")
    call alterar_campo("Password", senha, "ID")
    driver.FindElementByCss(".buttonstylenormal").Click
    Application.Wait (Now + TimeValue("0:00:02"))
    End Sub
' Fazer Login no JDE

Sub fechar_Chrome()
    driver.Quit
    Set driver = Nothing
    End Sub
' Fechar o chrome e encerrar o uso do driver

Function alterar_campo(campo As String, dado_inputado As String, tipo_proc_element as string)
    Select case tipo_proc_element
    case "ID"
        With driver
            .FindElementById(campo, timeout:=10000).Clear
            .ExecuteScript ("document.getElementById('" & campo & "').value = '" & dado_inputado & "'")
            .FindElementById(campo, timeout:=10000).SendKeys Enter
        End With
    case "Name"
        With driver
            .FindElementByName(campo, timeout:=10000).Clear
            .ExecuteScript ("document.getElementByName('" & campo & "').value = '" & dado_inputado & "'")
            .FindElementByName(campo, timeout:=10000).SendKeys Enter
        End With
    case Else
        debug.print "Mandou algo de errado aí no" & campo
    end Select
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
' Esperar carregar uma tela em até 10 segundos

Function Abrir_tela_fav(texto as string)
    With driver
            .FindElementById("drop_fav_menus").Click
            .FindElementByXPath("//div[3]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td/span").Click
            .FindElementByLinkText(texto).Click
    End With
    Application.Wait (Now + TimeValue("0:00:04"))        
    driver.SwitchToFrame 8
    end Function
' Selecionar uma tela do menu favoritos e abrir. Necessário melhor maneira de esperar do que Aplication Wait

Sub carregar_Exportar_JDE()
    Dim qtd_err As Integer
    qtd_err = 0
    Application.Wait (Now + TimeValue("0:00:05"))
    
    On Error Resume Next
    driver.FindElementById("jdehtmlGridLast0_1").Click

    Looping_carregar_arquivos:
    If qtd_err < 5 Then
        'Aguardar carregar - Não definido o valor de segundos correto ainda
        qtd_err = qtd_err + 1
        Application.Wait (Now + TimeValue("0:00:08"))

        'Exportar, ideal seria um código para checar
        With driver
        On Error GoTo Looping_carregar_arquivos
            .FindElementById("jdehtmlExportData0_1").Click
            .FindElementById("hc1").Click
        End With
        On Error GoTo 0
    Else
        Exit Sub
    End If

    ThisWorkbook.Activate
    Sheets("Temp").Activate
    Call pull_Book1xls

    Application.DisplayAlerts = False
    Range(Range("A1"), ActiveCell.SpecialCells(xlLastCell)).ClearContents
    Application.Wait (Now + TimeValue("0:00:15"))
    Workbooks.Open ("D:\Users\sb048948\Downloads\Book1.xlsx")

    Range(Range("A1"), ActiveCell.SpecialCells(xlLastCell)).Copy

    ThisWorkbook.Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Workbooks("Book1.xlsx").Close SaveChanges:=False
    Kill ("D:\Users\sb048948\Downloads\Book1.xlsx")
    Application.DisplayAlerts = True

    End Sub
' Esperar para exportar. Este código funciona, mas está longe de ser estável

Sub copiar_Temp_para_Pedidos()
    Sheets("Temp").Activate
    Range(Range("A2"), ActiveCell.SpecialCells(xlLastCell)).Copy
    Sheets("Pedidos Emitidos JDE").Activate
    Range("E2").End(xlDown).Select
    Selection.Offset(1, -4).Select
    ActiveSheet.Paste
    End Sub
' Copiar da planilha temp, após o Book1.xlsx e colar na planilha pedidos emitidos JDE

