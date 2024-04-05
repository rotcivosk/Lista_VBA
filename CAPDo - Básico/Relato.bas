Public session

Public WShell



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

Sub exportar_clipboardSAP(Is_SE16N As Boolean)
    With session
        If Is_SE16N Then
            .findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
            .findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem "&PC"
        Else
            .findById("wnd[0]/tbar[1]/btn[45]").press
        End If
        .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
        .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
        .findById("wnd[1]/tbar[0]/btn[0]").press
    End With
End Sub

Function barra_status(atual As Integer, total As Integer)
     
    DoEvents
    Application.StatusBar = "Carregando " & atual & " de " & total
    
    barra_status = 0
    
End Function

Function Abrir_popup_Windows(Texto As String) As Boolean
    'Achar o pop up - Imprimir
    Dim n As Integer
    Dim achou_tela
    Set WShell = CreateObject("WScript.Shell")
    n = 1
    Do
        achou_tela = WShell.AppActivate(Texto)
        n = n + 1
        Application.Wait Now + TimeValue("0:00:01")
    Loop Until achou_tela Or n > 10
    achou_tela = WShell.AppActivate(Texto)
        
    Abrir_popup_Windowns = achou_tela
        
End Function

Sub export_Formatar_PlanilhaSAP()
        
        
        Range(Range("A1"), ActiveCell.SpecialCells(xlLastCell)).ClearContents
        Range("A1").PasteSpecial
        Application.CutCopyMode = False
        With Columns("A:A")
            .TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar:="|", TrailingMinusNumbers:=True
            .Delete Shift:=xlToLeft
        
        Rows("1:3").Delete Shift:=xlUp
        Rows("2:2").Delete Shift:=xlUp
        End With
        
        Columns("A:A").Select
        With Range(Selection, Selection.End(xlToRight))
            .EntireColumn.AutoFit
            .NumberFormat = "General"
        End With
        
End Sub




