function limparPlan (nome as string)
    ThisWorkbook.Activate
    Sheets(nome).Activate
    Range(Range("A1"), ActiveCell.SpecialCells(xlLastCell)).ClearContents
end function