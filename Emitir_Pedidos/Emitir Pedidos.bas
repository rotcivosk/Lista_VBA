Sub emitir_pedidos()

    '***_____BASE____****
   
    call Abrir_SAP

    call criar_cotacao

    call criar_pedido

    
    ' Variáveis de Anexos
    Dim caminho_vbs As String, caminho_anexos As String
    caminho_anexos = "D:\Users\sb048948\OneDrive - Honda\Documentos\SAP\SAP GUI\"
    caminho_vbs = "D:\Users\sb048948\Downloads\Emitir_pedidos\"
    
    ' Nomes
    Dim proposta As String
    proposta = Range("d41").Value

    ' Seleciona o range que vai trabalhar e o range de anexos
    If Range("F41") = "" Then
        Set range_anexos = Range("F40")
    Else
        Set range_anexos = Range(Range("F40"), Range("F40").End(xlDown))
    End If
    

    ' Identador
    Dim ident As Integer
    ident = 0

    
End Sub

