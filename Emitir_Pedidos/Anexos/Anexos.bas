Function adicionar_anexos(caminho_anexos As String, caminho_vbs As String, descricao_PDF As String, eh_contrato As Boolean, precisa_pasta As Boolean)
       
    ' Apaga caso já exista o TxT
    Dim objeto2
    Set objeto2 = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    objeto2.DeleteFile (caminho_vbs & "pedido.txt")
    objeto2.DeleteFile (caminho_vbs & "flag.txt")
    On Error GoTo 0
    Set objeto2 = Nothing
    
    
    ' Grava um arquivo temporário com o nome do PDF
    Dim fso As Object
    Dim tempFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set tempFile = fso.CreateTextFile(caminho_vbs & "pedido.txt", True)
    If precisa_pasta Then
        tempFile.Write caminho_anexos & descricao_PDF
    Else
        tempFile.Write descricao_PDF
    End If
    tempFile.Close


    'Abrir e rodar Anexos
    With session
    
        'Abre o anexo
        .findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
        .findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_PCATTA_CREA"
        

        
        If eh_contrato Then
            .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
        
        Else
            'Seleciona cotação normal
            .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
        End If
    End With


    'Rodar macro.vbs
    Dim WScript, WSHShell
    Set WShell = CreateObject("WScript.Shell")
    WShell.Run caminho_vbs & "anexar_pdf.vbs", 1, False
    
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    ' Espera o .vbs terminar checando se há um arquivo temp
    Dim objfso
    Set objfso = CreateObject("Scripting.FileSystemObject")
    Do While Not objfso.FileExists(caminho_vbs & "flag.txt")
        Application.Wait Now + TimeValue("0:00:01")
    Loop

    ' Limpa tudo e marca como OK
    fso.DeleteFile (caminho_vbs & "pedido.txt")
    objfso.DeleteFile (caminho_vbs & "flag.txt")
    
End Function