Sub Adicionar_anexos_da_lista()
    
    'Variáveis de Anexos
    Dim caminho_vbs As String, caminho_anexos As String
    caminho_anexos = "D:\Users\sb048948\OneDrive - Honda\Documentos\SAP\SAP GUI\"
    caminho_vbs = "D:\Users\sb048948\Downloads\Emitir_pedidos\"

    'Abrir o SAP
    Dim Applic, Connection, SapGuiAuto
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Applic = SapGuiAuto.GetScriptingEngine
    Set Connection = Applic.Children(0)
    Set session = Connection.Children(0) 'Declara a Session pública como o SAP em aberto

    'Seleciona o range que vai trabalhar e o range de anexos
    If Range("F41") = "" Then
        Set range_anexos = Range("F40")
    Else
        Set range_anexos = Range(Range("F40"), Range("F40").End(xlDown))
    End If

    Dim temp As Boolean
    

    '***___CRIAR_ANEXOS___***
    For Each cell In range_anexos
        If cell.Offset(0, 1).Value = "Contrato" Then temp = True Else temp = False
        Call adicionar_anexos(caminho_anexos, caminho_vbs, cell.Value, temp, False)
    Next

End Sub

Function checar_saplmegui()
    
    Dim nome_comp As String
            
    Set user_megui = session.findById("wnd[0]/usr")
    i = 0
    For i = 0 To user_megui.Children.Count - 1
        nome_comp = user_megui.Children(CInt(i)).Name
        If Left(nome_comp, 15) = "SUB0:SAPLMEGUI:" Then
            Exit For
        End If
    Next
    
    checar_saplmegui = Right(nome_comp, 4)
    
End Function
