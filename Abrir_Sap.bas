Public session
Sub Abrir_SAP()

    Dim Applic, Connection, SapGuiAuto, WScript, WSHShell
    
    'Setar os Iniciais
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Applic = SapGuiAuto.GetScriptingEngine
    
    If Applic.Connections.Count() > 0 Then 'Checa se tem algum SAP em aberto
        Set Connection = Applic.Children(0)
        Set session = Connection.Children(0) 'Declara a Session pública como o SAP em aberto
    Else
        'Inputs do Usuário/Senha
        user = InputBox("Digite seu Usuário:", "Incluir User", " ")
        Senha = InputBox("Digite sua Senha:", "Incluir Senha", " ")

        'Abrir o arquivo SAP
        Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus
        Set WSHShell = CreateObject("WScript.Shell")
            Do Until WSHShell.AppActivate("SAP Logon ")
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        Set WSHShell = Nothing
        'Declara a Session pública como o novo SAP aberto
        Set Connection = Applic.OpenConnection("* 61 - ECP - Produção (001)", True)
        Set session = Connection.Children(0)

        'Logon
        session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = user
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Senha
        session.findById("wnd[0]").sendVKey 0
    End If

End Sub
