Public driver
'Lembre-se de add o Selenium Type Library em Ferramentas -> Referencias
Sub Abrir_Chrome(site As String)

    Set driver = New ChromeDriver
    driver.Get site
    driver.Window.maximize

End Sub

Sub Login_jde(user As String, Senha As String)
    Application.Wait (Now + TimeValue("0:00:05"))
    With driver
        .FindElementById("User").SendKeys [user]
        .FindElementById("Password").SendKeys [Senha]
        .FindElementByCss(".buttonstylenormal").Click
    End With
    Application.Wait (Now + TimeValue("0:00:05"))
End Sub

Sub fechar_Chrome()

    driver.Quit
    Set driver = Nothing
End Sub