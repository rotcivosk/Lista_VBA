Attribute VB_Name = "M_SAP_Abrir_Sap"
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
        'Inputs do Usuario/Senha
        user = InputBox("Digite seu Usuario:", "Incluir User", " ")
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
' Abrir o SAP ou usar o que ja esta ok


Sub exportar_clipboardSAP(Is_tabela As Boolean)

    'Ha duas maneiras de Exportar, uma caso seja uma tabela, outra caso seja uma transacao
    With session
        If Is_tabela Then
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
' Exportar o SAP selecionando o "Clipboard"

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
' Formatar a planilha para tirar as colunas