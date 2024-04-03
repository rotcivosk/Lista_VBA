Function ProcvArray(matMae As Variant, matFonte As Variant, numColunas As Integer, colunaInfo As Integer) As Variant
    Dim i As Long, j As Long, k As Integer
    Dim found As Boolean
    Dim result() As Variant

    ' Redimensiona o array resultante para incluir a coluna adicional
    ReDim result(1 To UBound(matMae, 1), 1 To UBound(matMae, 2) + 1)

    ' Copia os dados de matRequisicao para o array resultante
    For i = 1 To UBound(matMae, 1)
        For j = 1 To UBound(matMae, 2)
            result(i, j) = matMae(i, j)
        Next j
    Next i

    ' Realiza a operação de procura
    For i = 1 To UBound(matMae, 1)
        found = False
        For j = 1 To UBound(matFonte, 1)
            For k = 1 To numColunas
                If matMae(i, k) <> matFonte(j, k) Then
                    Exit For
                End If
            Next k
            If k > numColunas Then
                result(i, UBound(matMae, 2) + 1) = matFonte(j, colunaInfo)
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            result(i, UBound(matMae, 2) + 1) = "Não encontrado"
        End If
    Next i

    ProcvArray = result
End Function

