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