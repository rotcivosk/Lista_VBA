Function esperar_carregar(texto As String) As Boolean
    n = 0
    Do
        titulo_pag = driver.FindElementById("jdeFormTitle0", timeout:=10000).Text
        n = n + 1
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop Until titulo_pag = texto Or n > 10

    esperar_carregar = False
    If n > 9 Then
'        MsgBox ("Deu BO")
        esperar_carregar = True
    End If
End Function

Function Abrir_tela_fav(texto as string)
    With driver
            .FindElementById("drop_fav_menus").Click
            .FindElementByXPath("//div[3]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td/span").Click
            .FindElementByLinkText(texto).Click
    End With
END Function