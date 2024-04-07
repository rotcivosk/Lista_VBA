Sub clicar_botao_cotacao()


    For i = 1 To 15000
        myrow = Right(Space(10) & CStr(i), 11)
        On Error GoTo deu_erro
        session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").selectItem myrow, "&Hierarchy"
        session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem myrow, "&Hierarchy"
        session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[0]").pressButton "COPY"
        Exit Sub
pulou:
    Next
    On Error GoTo 0
    
    '
    '

    
Exit Sub

deu_erro:
Resume pulou

GoTo pulou

End Sub