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
