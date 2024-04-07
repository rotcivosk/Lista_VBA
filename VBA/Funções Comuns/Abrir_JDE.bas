Public driver

' Lembre-se de add o Selenium Type Library em Ferramentas -> Referencias
Sub Abrir_Chrome(site As String)

    Set driver = New ChromeDriver
    driver.Get site
    driver.Window.Maximize
  
End Sub

Sub Login_jde(user As String, Senha As String)

    With driver
        .FindElementById("User", timeout:=10000).SendKeys [user]
        .FindElementById("Password", timeout:=10000).SendKeys [Senha]
        .FindElementByCss(".buttonstylenormal", timeout:=10000).Click
    End With
End Sub

Sub fechar_Chrome()

    driver.Quit
    Set driver = Nothing
End Sub

