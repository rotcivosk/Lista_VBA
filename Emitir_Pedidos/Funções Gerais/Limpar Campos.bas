Sub Limpar_Campos()
'
' Limpar_Campos Macro
'

'
    With Sheets("Macro - Pedidos")
        .Activate
        .Range("B21:H36").ClearContents
        .Range("F40:H62").ClearContents
        .Range("D40").ClearContents
        .Range("G6:H7").ClearContents
    End With
    
    With Sheets("Temp")
        .Activate
        .Range(Range("A1"), ActiveCell.SpecialCells(xlLastCell)).ClearContents
    End With
    
End Sub