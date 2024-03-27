Sub Parte_1_SAP_EKPO_ME5A()
'
' Macro3 Macro
' reafdaeq1
'

'
    Dim lekpo As Long
    Dim lme5a As Long
    Dim leab As Long
    Dim Lme2n As Long

'ekpo
    
    Sheets("F - EKPO").Select
    lekpo = Range("A" & Rows.Count).End(xlUp).Row
    Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2:A" & lekpo).FormulaR1C1 = "=RC[8]&RC[9]"
    Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B2:B" & lekpo).FormulaR1C1 = "=RC[1]&RC[2]"
   
'me5a
    
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "REQ - SAP"
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "F - Temp"
    Sheets("F - ME5A").Select
    lme5a = Range("B" & Rows.Count).End(xlUp).Row
    Range("A1").FormulaR1C1 = "Index"
    Range("A2:A" & lme5a).FormulaR1C1 = "=RC[1]&RC[2]"
    Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D2:D" & lme5a).FormulaR1C1 = "=if(iserror(VLOOKUP(RC[-3],'F - EKPO'!C1:C3,3,FALSE)),""Em Aberto"",""Com Pedido"")"
    Rows("1:1").AutoFilter
    ActiveSheet.Range("$A$1:$Z$" & lme5a).AutoFilter Field:=4, Criteria1:="Em Aberto"
    Cells.Select
    Selection.Copy
    Sheets("REQ - SAP").Select
    Cells.Select
    ActiveSheet.Paste
    lme5a = Range("A" & Rows.Count).End(xlUp).Row
    
        

'me2n
    
    Sheets("Ped - SAP").Select
    Lme2n = Range("A" & Rows.Count).End(xlUp).Row
    
    Columns("I:I").Replace What:=".", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("F1").FormulaR1C1 = "ReqAprov"
    'Range("F2:F" & lme2n).FormulaR1C1 = ""
    
       
    Range("E:E,P:P").Replace What:=" ", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
    
    Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2:A" & Lme2n).FormulaR1C1 = "=value(RC[1]&RC[2])"
    Range("A1").FormulaR1C1 = "Índex"


'Puxar CDPOS e CDHDR com os valores de Req
    

    Sheets("Req - SAP").Select
    Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("F - Temp").Select
    Range("B1").Select
    ActiveSheet.Paste
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Requisição"
    
    Sheets("Ped - SAP").Select
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("F - Temp").Select
    leab = Range("B" & Rows.Count).End(xlUp).Row
    leab = leab + 1
    Range("B" & leab).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Columns("B:B").NumberFormat = "0000000000"
    ActiveSheet.Range("B:B").RemoveDuplicates Columns:=1, Header:=xlNo
    
'Puxar EORD, EINA e EINE com os valores de Materiais

    Sheets("Req - SAP").Select
    Range("G1:G" & lme5a).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("F - Temp").Select
    Range("D2").Select
    ActiveSheet.Paste
    Range("D1").FormulaR1C1 = "Materiais Geral"
    ActiveSheet.Range("$D$2:$D$" & lme5a).RemoveDuplicates Columns:=1, Header:=xlNo
    Range("D2:D" & lme5a).Select
    ActiveWorkbook.Worksheets("F - TEMP").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("F - TEMP").Sort.SortFields.Add Key:=Range("D2:D" & lme5a), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("F - TEMP").Sort
        .SetRange Range("D:D")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


    Sheets("F - ME5A").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("F - EKPO").Select
    ActiveWindow.SelectedSheets.Delete
    
    
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "F - CDHDR"
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "F - CDPOS"
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "F - EORD"
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "F - EINA"
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "F - EINE"
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "F - EKKO"
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "Relato Semana Anterior"
    
End Sub
Sub Parte_2_SAP_INFO()
'
' Segunda_Parte Macro
' Continuação do relatório de requisição em Aberto
'

'
    Dim Ltemp As Long
    Dim Lme2n As Long
    

' Organizar CDHDR

    Sheets("F - CDHDR").Select
    Ltemp = Range("A" & Rows.Count).End(xlUp).Row
    
    
' Teste das datas
    
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Columns("E:E").Select
    Selection.TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=".", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Range("H2:H" & Ltemp).FormulaR1C1 = "=DATE(RC[-1],RC[-2],RC[-3])"
    Columns("H:H").EntireColumn.AutoFit
    Range("E1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H1").Select
    ActiveSheet.Paste
    Columns("H:H").Select
    Application.CutCopyMode = False
    Selection.Copy
    Columns("E:E").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:H").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
    Range("A1:M1").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("F - CDHDR").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("F - CDHDR").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "E1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("F - CDHDR").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
' Procurar Datas

    Sheets("Req - SAP").Select
    Ltemp = Range("A" & Rows.Count).End(xlUp).Row
    Columns("D:D").Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").FormulaR1C1 = "Data de Aprovação"
    Range("E2:E" & Ltemp).FormulaR1C1 = "=VLOOKUP(RC[-3],'F - CDHDR'!C[-3]:C[1],4,FALSE)"
    Range("F1").FormulaR1C1 = "Dias em Aberto"
    Range("F2:F" & Ltemp).FormulaR1C1 = "=NETWORKDAYS(RC[-1],TODAY())"
    Columns("E:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Sheets("Ped - SAP").Select
    Lme2n = Range("A" & Rows.Count).End(xlUp).Row
    Range("G2:G" & Lme2n).FormulaR1C1 = "=VLOOKUP(RC[-3],'F - CDHDR'!C[-5]:C[-2],4,FALSE)"
    Columns("G:G").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
'Organizar EORD, EINA e EINE
   
    Sheets("F - EORD").Select
    Ltemp = Range("A" & Rows.Count).End(xlUp).Row
    Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2:A" & Ltemp).FormulaR1C1 = "=RC[1]&RC[2]"
        
    Sheets("F - EINA").Select
    Ltemp = Range("A" & Rows.Count).End(xlUp).Row
    Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2:A" & Ltemp).FormulaR1C1 = "=RC[2]&RC[4]"
        
    Sheets("F - EINE").Select
    Ltemp = Range("A" & Rows.Count).End(xlUp).Row
    Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2:A" & Ltemp).FormulaR1C1 = "=RC[1]&RC[4]"
    

'Inserir Colunas para RegInfo
    
    Sheets("Req - SAP").Select
    Ltemp = Range("A" & Rows.Count).End(xlUp).Row
    Columns("J:J").Delete Shift:=xlToLeft
    Columns("S:S").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("S1").FormulaR1C1 = "Fornecedor"
    Range("S2:S" & Ltemp).FormulaR1C1 = "=VLOOKUP(RC[-11]&RC[-2],'F - EORD'!C[-18]:C[-10],9,FALSE)"
    Range("T1").FormulaR1C1 = "RegInfo"
    Range("T2:T" & Ltemp).FormulaR1C1 = "=VLOOKUP(RC[-12]&RC[-1],'F - EINA'!C[-19]:C[-18],2,FALSE)"
    Range("U1").FormulaR1C1 = "Cancelado LOF"
    Range("U2:U" & Ltemp).FormulaR1C1 = "=VLOOKUP(RC[-13]&RC[-2],'F - EINA'!C[-20]:C[-15],6,FALSE)"
    Range("V1").FormulaR1C1 = "Cancelado Centro"
    Range("V2:V" & Ltemp).FormulaR1C1 = "=VLOOKUP(RC[-2]&RC[-5],'F - EINE'!C[-21]:C[-16],6,FALSE)"
       
    
'Inserir Colunas Que faltam

    
'Tirar #N/D

    Columns("S:V").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("S2:Y" & Ltemp).Select
    Selection.Replace What:="#N/A", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'Tirar o espaço

    Range("S2:Y" & Ltemp).Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


'Remendos
    
    Columns("S:S").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("S1").FormulaR1C1 = "Tipo"
    Range("S2:S" & Ltemp).FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(VALUE(RC[-18]),'Relato Semana Anterior'!C[-18],1,FALSE)),IF(RC[-13]>180,""Erro de Sistema"",IF(AND(RC[3]="""",RC[4]="""",NOT(RC[1]="""")),""RegInfo"",IF(RC[-12]=""A"",IF(RC[-11]="""",""Investimento Serv"",""Investimento Mat""),IF(RC[-11]="""",""Separar Serviço e Contrato"",""Material"")))),VLOOKUP(VALUE(RC[-18]),'Relato Semana Anterior'!C[-18]:C[1],20,FALSE))"
        
     
    Columns("Q:Q").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("Q1").FormulaR1C1 = "Valor Cabeçalho"
    Range("Q2:Q" & Ltemp).FormulaR1C1 = "=SUMIF(C[-15],RC[-15],C[-1])"
    
    Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G1").FormulaR1C1 = "Comprador"
    Range("G2:G" & Ltemp).FormulaR1C1 = "=IF(RC[14]=""MATERIAL"",VLOOKUP(RC[8]&VLOOKUP(RC[9],'CÁLCULOS BASE'!C[9]:C[14],2,FALSE),'CÁLCULOS BASE'!C[2]:C[3],2,FALSE),IF(RC[14]=""SERVIÇO"",IF(RC[11]<5000,""MICHELE"",""LUCIO""),VLOOKUP(RC[14],'CÁLCULOS BASE'!C[5]:C[7],3,FALSE)))"
    

    Columns("T:T").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("T1").FormulaR1C1 = "Data de Remessa"
    Range("T2:T" & Ltemp).FormulaR1C1 = "=DATE(MID(RC[1],1,4),MID(RC[1],5,2),MID(RC[1],7,2))"
    Columns("T:T").Copy
    Columns("T:T").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    Columns("U:U").Delete Shift:=xlToLeft
    



    Sheets("Ped - SAP").Select
    Lme2n = Range("A" & Rows.Count).End(xlUp).Row
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("D2:F" & Lme2n).FormulaR1C1 = "=VALUE(RC[-3])"
    Range("D2:F" & Lme2n).Copy
    Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Columns("D:F").Delete Shift:=xlToLeft
    
    
    Sheets("F - EORD").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("F - EINA").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("F - EINE").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("F - CDPOS").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("F - CDHDR").Select
    ActiveWindow.SelectedSheets.Delete
    
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "COT - ANTERIOR"
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "COT - JDE"
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "F - APROV"
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "REQ - JDE"
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "PED - JDE"
    

    Sheets("Req - Sap").Select
    Range("A:A").Copy
    Range("A:A").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    

    Sheets("Relato Semana Anterior").Select
    Range("A:A").Copy
    Range("A:A").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    
   
    End Sub
Sub Parte_3()
'
     Dim laprov As Long
    
    Sheets("F - APROV").Select
    laprov = Range("A" & Rows.Count).End(xlUp).Row
    Columns("H:H").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("G2"), DataType:=xlFixedWidth, _
        OtherChar:=".", FieldInfo:=Array(Array(0, 1), Array(4, 1), Array(6, 1)), _
        TrailingMinusNumbers:=True
    Range("J2:J" & laprov).FormulaR1C1 = "=DATE(RC[-3],RC[-2],RC[-1])"
    Range("J2:J" & laprov).Copy
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("H:J").Select
    Range("J1").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("A1:I1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("F - APROV").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("F - APROV").Sort.SortFields.Add Key:=Range( _
        "G2:G" & laprov), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("F - APROV").Sort
        .SetRange Range("A1:I" & laprov)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   
    Sheets("REQ - JDE").Select
    laprov = Range("A" & Rows.Count).End(xlUp).Row
    Columns("G:G").Cut
    Columns("A:A").Insert Shift:=xlToRight
    Columns("W:W").Cut
    Columns("B:B").Insert Shift:=xlToRight
    Range("A1").FormulaR1C1 = "Requisição"
    Range("B1").FormulaR1C1 = "Linha"
    Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C2:C" & laprov).FormulaR1C1 = "=RC[-1]/100"
    Range("C2:C" & laprov).Select
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("C2:C" & laprov).FormulaR1C1 = "=RC[16]/100"
    Range("C2:C" & laprov).Copy
    Range("S2").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    Range("C2:C" & laprov).FormulaR1C1 = "=RC[18]/100"
    Range("C2:C" & laprov).Copy
    Range("U2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("C2:C" & laprov).FormulaR1C1 = "=RC[17]/10000"
    Range("C2:C" & laprov).Copy
    Range("T2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Columns("C:C").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
    
    Columns("H:H").Cut
    Columns("C:C").Insert Shift:=xlToRight
    Columns("N:N").Cut
    Columns("D:D").Insert Shift:=xlToRight
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E2:E" & laprov).FormulaR1C1 = "=VALUE(RC[-1])"
    Range("E2:E" & laprov).Copy
    Range("D2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Range("E2:E" & laprov).FormulaR1C1 = "=VLOOKUP(RC[-4],'F - APROV'!C[-3]:C[2],6,FALSE)"

    Range("E1").FormulaR1C1 = "Data de Aprovação"
    Range("D1").FormulaR1C1 = "Data de Emissão"
    Range("C1").FormulaR1C1 = "Tipo do Pedido"
    Columns("N:O").Cut
    Columns("F:F").Insert Shift:=xlToRight
    Range("F1").FormulaR1C1 = "Código do Material"
    Range("G1").FormulaR1C1 = "Descrição do Material"
    Range("K1").FormulaR1C1 = "Últ. Status"
    Range("L1").FormulaR1C1 = "Pró. Status"
    Columns("K:L").Cut
    Columns("C:C").Insert Shift:=xlToRight
    Columns("F:G").NumberFormat = "m/d/yyyy"
    Columns("S:U").Cut
    Columns("J:J").Insert Shift:=xlToRight
    Columns("X:X").Cut
    Columns("M:M").Insert Shift:=xlToRight
    Columns("Q:Q").Cut
    Columns("N:N").Insert Shift:=xlToRight
    Columns("P:AD").Delete Shift:=xlToLeft
    Range("J1").FormulaR1C1 = "Quantidade do Pedido"
    Range("K1").FormulaR1C1 = "Valor Unitário"
    Range("L1").FormulaR1C1 = "Valor Total"
    Range("M1").FormulaR1C1 = "Unidade de Medida"
    Range("N1").FormulaR1C1 = "Requisitante"
    Range("O1").FormulaR1C1 = "Filial"
    Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H1").FormulaR1C1 = "Tipo"
    Columns("I:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("I1").FormulaR1C1 = "Comprador"
    Range("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
   
    Range("I2:I" & laprov).FormulaR1C1 = "=IF(RC[-3]=""OY"",IF(MID(RC[2],1,1)=""Q"",""Importado Q"",""Importado Desp""),IF(MID(RC[2],1,1)=""Q"",""Investimento"",IF(ISERROR(VLOOKUP(RC[3],'CÁLCULOS BASE'!C[-4],1,FALSE)),""Material"",""Contrato"")))"
    Range("J2:J" & laprov).FormulaR1C1 = "=IF(RC[-1]=""Material"",VLOOKUP(MID(RC[1],1,2),'CÁLCULOS BASE'!C[-8]:C[-5],2,FALSE),VLOOKUP(RC[-1],'CÁLCULOS BASE'!C[2]:C[6],3,FALSE))"
    Range("H2:H" & laprov).FormulaR1C1 = "=VLOOKUP(RC[-6],'F - APROV'!C[-6]:C[1],6,FALSE)"
    Range("A2:A" & laprov).FormulaR1C1 = "=Value(RC[1]&RC[2])"
    Range("A1").FormulaR1C1 = "Índex"
    
        Dim lcot As Long
    
    Sheets("COT - JDE").Select
    lcot = Range("A" & Rows.Count).End(xlUp).Row
'
    Columns("G:G").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:I").Select
    Range("I1").Delete Shift:=xlToLeft
    Range("A1").FormulaR1C1 = "Filial"
    Range("B1").FormulaR1C1 = "Status"
    Range("C1").FormulaR1C1 = "Requisição"
    Range("D1").FormulaR1C1 = "Ult Status"
    Range("E1").FormulaR1C1 = "Pro Status"
    Range("F1").FormulaR1C1 = "Tipo"
    Range("G1").FormulaR1C1 = "Comprador"
    Range("H1").FormulaR1C1 = "Tipo"
    Range("I1").FormulaR1C1 = "Req Aprov"
    Range("J1").FormulaR1C1 = "Núm Pedido"
    Range("K1").FormulaR1C1 = "Comprador"
    Range("L1").FormulaR1C1 = "Cód Fornecedor"
    Range("M1").FormulaR1C1 = "Forncedor"
    Range("N1").FormulaR1C1 = "Cód. Item"
    Range("O1").FormulaR1C1 = "Descrição Item"
    Range("P1").FormulaR1C1 = "Data Pedido"
    Range("Q1").FormulaR1C1 = "Data Recebm"
    Range("R1").FormulaR1C1 = "Data de Remessa"
    Range("S1").FormulaR1C1 = "N° NF"
    Range("T1").FormulaR1C1 = "Quantidade"
    Columns("T:T").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("T2:T" & lcot).FormulaR1C1 = "=RC[3]/100"
    Range("U2:U" & lcot).FormulaR1C1 = "=RC[3]/10000"
    Range("V2:V" & lcot).FormulaR1C1 = "=RC[3]/100"


    Range("T2:V" & lcot).Copy
    Range("W2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Columns("T:V").Delete Shift:=xlToLeft
    Range("U1").FormulaR1C1 = "Vlr Unit"
    Range("V1").FormulaR1C1 = "Vlr Ped"
    Range("W1").FormulaR1C1 = "Tipo"
    Range("X1").FormulaR1C1 = "Requisição"
    Range("Y1").FormulaR1C1 = "UM"
    Range("Z1").FormulaR1C1 = "Linha"
    
    Range("B2:B" & lcot).FormulaR1C1 = "=IF(ISERROR(VLOOKUP(RC[1],'COT - ANTERIOR'!C[1],1,FALSE)),""Nova"",""Antiga"")"
    Range("C2:C" & lcot).FormulaR1C1 = "=VALUE(RC[21])"
    Range("G2:G" & lcot).FormulaR1C1 = "=VLOOKUP(RC[4],'CÁLCULOS BASE'!C[5]:C[7],2,FALSE)"
    Range("H2:H" & lcot).FormulaR1C1 = "=IF(RC[15]=""OY"",IF(MID(RC[7],1,1)=""Q"",""Importado Q"",""Importado Desp""),IF(RC[-1]=""Michele"",""Contrato"",IF(MID(RC[6],1,1)=""Q"",""Investimento"",""Material"")))"
    Range("I2:I" & lcot).FormulaR1C1 = "=VLOOKUP(RC[-6],'F - APROV'!C[-7]:C[-2],6,FALSE)"
    Columns("J:J").Insert Shift:=xlToRight
    Range("J2:J" & lcot).FormulaR1C1 = "=IF(NETWORKDAYS(RC[-1],TODAY())>VLOOKUP(RC[-2],'CÁLCULOS BASE'!C[2]:C[4],2,FALSE),""Fora do Prazo"",""Dentro do Prazo"")"
    
    
    Sheets("PED - JDE").Select
    lcot = Range("E" & Rows.Count).End(xlUp).Row
    Range("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A2:A" & lcot).FormulaR1C1 = "=Value(RC[5]&RC[6])"
    Range("A1").FormulaR1C1 = "Index"
    
    
End Sub

Sub Parte_4_Formulas()
'
' Formulas Macro
'
    Dim ltotal As Long
    Dim lsap As Long
    Dim ljde As Long
    
    
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "Requisição - Consolidado"

 
    Sheets("Req - SAP").Select
    lsap = Range("A" & Rows.Count).End(xlUp).Row
    Range("A2:A" & lsap).Copy
    
    Sheets("Requisição - Consolidado").Select
    Range("A2").Select
    ActiveSheet.Paste
    
    Sheets("Req - JDE").Select
    ljde = Range("A" & Rows.Count).End(xlUp).Row
    Range("A2:A" & ljde).Copy
    lsap = lsap + 1
    Sheets("Requisição - Consolidado").Select
    Range("A" & lsap).Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    ltotal = lsap + ljde - 2
       


'Fórmulas Requisição em aberto

    Range("A1").FormulaR1C1 = "Req + Item"
    Range("B1").FormulaR1C1 = "Requisição"
    Range("C1").FormulaR1C1 = "Item"
    Range("D1").FormulaR1C1 = "Origem"
    Range("E1").FormulaR1C1 = "TdDc"
    Range("F1").FormulaR1C1 = "C"
    Range("G1").FormulaR1C1 = "Comprador"
    Range("H1").FormulaR1C1 = "Data Liberação"
    Range("I1").FormulaR1C1 = "Dias em Aberto"
    Range("J1").FormulaR1C1 = "Status"
    Range("K1").FormulaR1C1 = "Fornecedor"
    Range("L1").FormulaR1C1 = "ReqInfo"
    Range("M1").FormulaR1C1 = "Material"
    Range("N1").FormulaR1C1 = "Texto"
    Range("O1").FormulaR1C1 = "Grupo de Mercadoria"
    Range("P1").FormulaR1C1 = "Tp. De Aval"
    Range("Q1").FormulaR1C1 = "Valor Requisição"
    Range("R1").FormulaR1C1 = "Lead Time"
    Range("S1").FormulaR1C1 = "CHECK GRAFICO"
    Range("T1").FormulaR1C1 = "Tipo"
    Range("U1").FormulaR1C1 = "Check"
    Range("V1").FormulaR1C1 = "Código de Linha"
    
    
    
    Range("B2:B" & ltotal).FormulaR1C1 = "=IF(RC[2]=""SAP"",MID(RC[-1],1,8),MID(RC[-1],1,7))"
    Range("C2:C" & ltotal).FormulaR1C1 = "=IF(RC[1]=""SAP"",MID(RC[-2],9,4),MID(RC[-2],8,4))"
    Range("D2:D" & ltotal).FormulaR1C1 = "=IF(ISERROR(VLOOKUP(RC[-3],'REQ - SAP'!C[-3],1,FALSE)),""JDE"",""SAP"")"
    Range("E2:E" & ltotal).FormulaR1C1 = "=IF(RC[-1]=""JDE"",VLOOKUP(RC[-4],'REQ - JDE'!C[-4]:C[1],6,FALSE),VLOOKUP(RC[-4],'REQ - SAP'!C[-4]:C[-1],4,FALSE))"
    Range("F2:F" & ltotal).FormulaR1C1 = "=IF(RC[-2]=""sap"",VLOOKUP(RC[-5],'REQ - SAP'!C[-5]:C[2],8,FALSE),"""")"
    Range("G2:G" & ltotal).FormulaR1C1 = "=IF(RC[-3]=""SAP"",VLOOKUP(RC[-6],'REQ - SAP'!C[-6]:C,7,FALSE),VLOOKUP(RC[-6],'REQ - JDE'!C[-6]:C[3],10,FALSE))"
    Range("H2:H" & ltotal).FormulaR1C1 = "=IF(RC[-4]=""SAP"",VLOOKUP(RC[-7],'REQ - SAP'!C[-7]:C[-3],5,FALSE),VLOOKUP(RC[-7],'REQ - JDE'!C[-7]:C,8,FALSE))"
    Range("I2:I" & ltotal).FormulaR1C1 = "=NETWORKDAYS(RC[-1],TODAY())-1"
    Range("J2:J" & ltotal).FormulaR1C1 = "=IF(RC[8]>RC[-1],""Dentro do Prazo"", ""Fora do Prazo"")"
    Range("K2:K" & ltotal).FormulaR1C1 = "=IF(RC[-7]=""SAP"",IF(VLOOKUP(RC[-10],'REQ - SAP'!C[-10]:C[14],24,FALSE)=0,"""",VLOOKUP(RC[-10],'REQ - SAP'!C[-10]:C[14],24,FALSE)),"""")"
    Range("L2:L" & ltotal).FormulaR1C1 = "=IF(RC[8]=""RegInfo"",IF(VLOOKUP(RC[-11],'REQ - SAP'!C[-11]:C[13],25,FALSE)=0,"""",VLOOKUP(RC[-11],'REQ - SAP'!C[-11]:C[13],25,FALSE)),"""")"
    Range("M2:M" & ltotal).FormulaR1C1 = "=IF(RC[-9]=""SAP"",VLOOKUP(RC[-12],'REQ - SAP'!C[-12]:C[-4],9,FALSE),VLOOKUP(RC[-12],'REQ - JDE'!C[-12]:C[-2],11,FALSE))"
    Range("N2:N" & ltotal).FormulaR1C1 = "=IF(RC[-10]=""SAP"",VLOOKUP(RC[-13],'REQ - SAP'!C1:C[-4],10,FALSE),VLOOKUP(RC[-13],'REQ - JDE'!C[-13]:C[-2],12,FALSE))"
    Range("O2:O" & ltotal).FormulaR1C1 = "=IF(RC[-11]=""SAP"",VLOOKUP(RC[-14],'REQ - SAP'!C[-14]:C[1],15,FALSE),"""")"
    Range("P2:P" & ltotal).FormulaR1C1 = "=IF(RC[-12]=""SAP"",IF(VLOOKUP(RC[-15],'REQ - SAP'!C[-15]:C[3],16,FALSE)=0,"""",VLOOKUP(RC[-15],'REQ - SAP'!C[-15]:C[3],16,FALSE)),"""")"
    Range("Q2:Q" & ltotal).FormulaR1C1 = "=IF(RC[-13]=""SAP"",VLOOKUP(RC[-16],'REQ - SAP'!C[-16]:C[2],18,FALSE),VLOOKUP(RC[-16],'REQ - JDE'!C[-16]:C[-2],15,FALSE))"
    Range("R2:R" & ltotal).FormulaR1C1 = "=VLOOKUP(RC[2],'CÁLCULOS BASE'!C[-6]:C[-5],2,FALSE)"
    Range("S2:S" & ltotal).FormulaR1C1 = "=IF(RC[-10]<5,""5"",IF(RC[-10]<10,""10"",IF(RC[-10]<20,""15"",""20"")))"
    Range("T2:T" & ltotal).FormulaR1C1 = "=IF(RC[-16]=""SAP"",VLOOKUP(RC[-19],'REQ - SAP'!C[-19]:C[3],21,FALSE),VLOOKUP(RC[-19],'REQ - JDE'!C[-19]:C[-11],9,FALSE))"
    Range("U2:U" & ltotal).FormulaR1C1 = "=IF(ISERROR(VLOOKUP(RC[-20],'F - TEMP'!C[-11],1,FALSE)),""Linha"",""Cabeçalho"")"
    Range("V1:V" & ltotal).FormulaR1C1 = "=IF(ISERROR(VLOOKUP(RC[-21],'Relato Semana Anterior'!C[-21],1,FALSE)),""Nova Requisição"",""Em Aberto"")"
    
    
    Columns("A:C").Copy
    Sheets("F - TEMP").Select
    Columns("J:J").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "Ped - Consolidado"
    
    Sheets("Ped - SAP").Select
    lsap = Range("A" & Rows.Count).End(xlUp).Row
    Range("A2:A" & lsap).Copy
    
    Sheets("Ped - Consolidado").Select
    Range("A2").Select
    ActiveSheet.Paste
    
    Sheets("Ped - JDE").Select
    ljde = Range("A" & Rows.Count).End(xlUp).Row
    Range("A2:A" & ljde).Copy
    
    lsap = lsap + 1
    
    Sheets("Ped - Consolidado").Select
    Range("A" & lsap).Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
    ltotal = lsap + ljde - 2
    Sheets("Ped - Consolidado").Select
    Range("A1").FormulaR1C1 = "Índex"
    Range("B1").FormulaR1C1 = "Pedido"
    Range("B2:B" & ltotal).FormulaR1C1 = "=IF(RC[6]=""JDE"",VLOOKUP(RC[-1],'PED - JDE'!C[-1]:C[5],6,FALSE),VLOOKUP(RC[-1],'PED - SAP'!C[-1]:C,2,FALSE))"
    Range("C1").FormulaR1C1 = "Linha Ped"
    Range("C2:C" & ltotal).FormulaR1C1 = "=IF(RC[5]=""JDE"",VLOOKUP(RC[-2],'PED - JDE'!C[-2]:C[5],7,FALSE),VLOOKUP(RC[-2],'PED - SAP'!C[-2]:C,3,FALSE))"
    Range("D1").FormulaR1C1 = "INDEX"
    Range("D2:D" & ltotal).FormulaR1C1 = "=VALUE(RC[1]&RC[2])"
    Range("E1").FormulaR1C1 = "Requisição"
    Range("E2:E" & ltotal).FormulaR1C1 = "=IF(RC[3]=""JDE"",IF(VLOOKUP(RC[-4],'PED - JDE'!C[-4]:C[-2],2,FALSE)=""0"",VLOOKUP(RC[-4],'PED - JDE'!C[-4]:C[0],4,FALSE),VLOOKUP(RC[-4],'PED - JDE'!C[-4]:C[-2],2,FALSE)),VLOOKUP(RC[-4],'PED - SAP'!C[-4]:C[-1],4,FALSE))"
    Range("F1").FormulaR1C1 = "Linha Req"
    Range("F2:F" & ltotal).FormulaR1C1 = "=IF(RC[2]=""JDE"",VLOOKUP(RC[-5],'PED - JDE'!C[-5]:C[-2],3,FALSE),VLOOKUP(RC[-5],'PED - SAP'!C[-5]:C[-1],5,FALSE))"
    Range("G1").FormulaR1C1 = "RegInfo"
    Range("G2:G" & ltotal).FormulaR1C1 = "=IF(RC[1]=""JDE"","""",VLOOKUP(RC[-6],'PED - SAP'!C[-6]:C[-1],6,FALSE))"
    Range("H1").FormulaR1C1 = "Origem"
    Range("H2:H" & ltotal).FormulaR1C1 = "=IF(ISERROR(VLOOKUP(RC[-7],'PED - SAP'!C[-7],1,FALSE)),""JDE"",""SAP"")"
    Range("I1").FormulaR1C1 = "Material"
    Range("I2:I" & ltotal).FormulaR1C1 = "=IF(RC[-1]=""JDE"",VLOOKUP(RC[-8],'PED - JDE'!C[-8]:C[12],20,FALSE),VLOOKUP(RC[-8],'PED - SAP'!C[-8]:C[17],13,FALSE))"
    Range("J1").FormulaR1C1 = "Descrição"
    Range("J2:J" & ltotal).FormulaR1C1 = "=IF(RC[-2]=""JDE"",VLOOKUP(RC[-9],'PED - JDE'!C[-9]:C[12],21,FALSE),VLOOKUP(RC[-9],'PED - SAP'!C[-9]:C[8],14,FALSE))"
    Range("K1").FormulaR1C1 = "Tipo"
    Range("K2:K" & ltotal).FormulaR1C1 = "=IF(RC[-3]=""JDE"",VLOOKUP(RC[-10],'PED - JDE'!C[-10]:C[2],12,FALSE),VLOOKUP(RC[-10],'PED - SAP'!C[-10]:C[2],8,FALSE))"
    Range("L1").FormulaR1C1 = "Comprador Original"
    Range("L2:L" & ltotal).FormulaR1C1 = "=IF(RC[8]=""Nova Requisição"",RC[1],VLOOKUP(RC[-8],'Relato Semana Anterior'!C[-11]:C[-5],7,FALSE))"
    Range("M1").FormulaR1C1 = "Emissão Pedido"
    Range("M2:M" & ltotal).FormulaR1C1 = "=VLOOKUP(IF(RC[-5]=""JDE"",VLOOKUP(RC[-12],'PED - JDE'!C[-12]:C[3],16,FALSE),VLOOKUP(RC[-11],'F - EKKO'!C[-12]:C[-4],9,FALSE)),'CÁLCULOS BASE'!C[-1]:C[1],2,FALSE)"
    Range("N1").FormulaR1C1 = "Tipo"
    Range("N2:N" & ltotal).FormulaR1C1 = "=IF(RC[-9]=0,""Frete"",IF(RC[-6]=""SAP"",IF(RC[-7]<>0,""RegInfo"",IF(RC[-3]=""A"",""Investimento"",IF(RC[-5]=0,""Serviço"",""Material""))),IF(RC[-3]=""OL"",""Importado Desp"",IF(RC[-3]=""OM"",""Importado Q"",IF(MID(RC[-4],1,1)=""Q"",""Investimento"",IF(RC[-1]=""Michele"",""Contrato"",""Material""))))))"
    Range("O1").FormulaR1C1 = "Aprovação Req"
    Range("O2:O" & ltotal).FormulaR1C1 = "=IF(RC[-7]=""JDE"",VLOOKUP(RC[-10],'F - APROV'!C[-13]:C[-8],6,FALSE),VLOOKUP(RC[-14],'PED - SAP'!C[-14]:C[-5],7,FALSE))"
    Range("P1").FormulaR1C1 = "Emissão Ped"
    Range("P2:P" & ltotal).FormulaR1C1 = "=IF(RC[-8]=""JDE"",VLOOKUP(RC[-15],'PED - JDE'!C[-15]:C[-1],15,FALSE),VLOOKUP(RC[-15],'PED - SAP'!C[-15]:C[-5],11,FALSE))"
    Range("Q1").FormulaR1C1 = "Lead Time"
    Range("Q2:Q" & ltotal).FormulaR1C1 = "=NETWORKDAYS(RC[-2],RC[-1])"
    Range("R1").FormulaR1C1 = "Status"
    Range("R2:R" & ltotal).FormulaR1C1 = "=IF(RC[-4]=""Frete"",""Frete"",IF(RC[-1]>VLOOKUP(RC[-4],'CÁLCULOS BASE'!C[-6]:C[-4],2,FALSE),""Fora do Prazo"",""Dentro do Prazo""))"
    Range("S1").FormulaR1C1 = "Comprador Original"
    Range("S2:S" & ltotal).FormulaR1C1 = "=IF(RC[-7]=RC[-6],""Mesmo"",""Diferente"")"
    Range("T1").FormulaR1C1 = "Status"
    Range("T2:T" & ltotal).FormulaR1C1 = "=IF(ISERROR(VLOOKUP(RC[-16],'Relato Semana Anterior'!C[-19],1,FALSE)),""Nova Requisição"",""Em aberto"")"
    
    Columns("I:I").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("I:I" & Ltemp).Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

        
    Sheets("Modelo CAPDo - Semanal").Select
    Range("D4").FormulaR1C1 = "=SUM(R[1]C:R[3]C)"
    Range("D5:D6").FormulaR1C1 = "=COUNTIFS('RELATO SEMANA ANTERIOR'!C[3],R1C11,'RELATO SEMANA ANTERIOR'!C[6],RC[-1])"
    Range("D7:D8").FormulaR1C1 = "=COUNTIFS('COT - ANTERIOR'!C[3],R1C11,'COT - ANTERIOR'!C[6],RC[-1])"
    Range("D10").FormulaR1C1 = "=SUM(R[1]C:R[2]C)"
    Range("D11").FormulaR1C1 = "=COUNTIFS('Requisição - CONSOLIDADO'!C[3],R1C11,'Requisição - CONSOLIDADO'!C[18],""Nova Requisição"")"
    Range("D12").FormulaR1C1 = "=COUNTIFS('ped - Consolidado'!C[8],R1C11,'ped - Consolidado'!C[16],""Nova Requisição"")"
    Range("D14").FormulaR1C1 = "=SUM(R[1]C:R[4]C)"
    Range("D15").FormulaR1C1 = "=COUNTIFS('PED - CONSOLIDADO'!C[15],""Mesmo"",'PED - CONSOLIDADO'!C[9],r1c11)"
    Range("D16").FormulaR1C1 = "=COUNTIFS('PED - CONSOLIDADO'!C[15],""Diferente"",'PED - CONSOLIDADO'!C[8],R1C11)"
    Range("D18").FormulaR1C1 = "=SUM(R[1]C:R[2]C)"
    Range("D19:D20").FormulaR1C1 = "=COUNTIFS('Requisição - Consolidado'!C[3],R1C11,'Requisição - Consolidado'!C[6],RC[-1])"
  

    Sheets("Gráfico").Select

    Range("J9:M12").FormulaR1C1 = "=COUNTIFS('Requisição - Consolidado'!C7,RC6,'Requisição - Consolidado'!C19,R8C,'Requisição - Consolidado'!C20,RC9,'Requisição - Consolidado'!C21,""Cabeçalho"")"
    Range("J13:M22").FormulaR1C1 = "=COUNTIFS('Requisição - Consolidado'!C7,RC6,'Requisição - Consolidado'!C19,R8C,'Requisição - Consolidado'!C20,RC9)"



    Range("N9:O22").FormulaR1C1 = "=COUNTIFS('COT - JDE'!C[-7],rc6,'COT - JDE'!C[-4],R8C)"
    Range("P9:P10").FormulaR1C1 = "=SUM(RC[-6]:R[1]C[-1])"
         
    Range("J23:M23").FormulaR1C1 = "=SUM(R[-13]C:R[-1]C)"
     

    Range("X9:Y18").FormulaR1C1 = "=COUNTIFS('PED - CONSOLIDADO'!C14,RC23,'PED - CONSOLIDADO'!C18,R7C)"

    
    
    





End Sub

