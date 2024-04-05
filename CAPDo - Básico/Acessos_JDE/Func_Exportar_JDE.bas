Sub carregar_Exportar_JDE()
Dim qtd_err As Integer
qtd_err = 0

    Application.Wait (Now + TimeValue("0:00:05"))
    
    On Error Resume Next
    driver.FindElementById("jdehtmlGridLast0_1").Click

Looping_carregar_arquivos_OPC:
    If qtd_err < 5 Then
        'Aguardar carregar - Não definido o valor de segundos correto ainda
        qtd_err = qtd_err + 1
        Application.Wait (Now + TimeValue("0:00:08"))

        'Exportar, ideal seria um código para checar
        With driver
        On Error GoTo Looping_carregar_arquivos_OPC
            .FindElementById("jdehtmlExportData0_1").Click
            .FindElementById("hc1").Click
        End With
        On Error GoTo 0
    Else
        Exit Sub
    End If

    ThisWorkbook.Activate
    Sheets("Temp").Activate
    Call pull_Book1xls
End Sub

Sub pull_Book1xls()

    Application.DisplayAlerts = False
    Range(Range("A1"), ActiveCell.SpecialCells(xlLastCell)).ClearContents
    Application.Wait (Now + TimeValue("0:00:15"))
    Workbooks.Open ("D:\Users\sb048948\Downloads\Book1.xlsx")

    Range(Range("A1"), ActiveCell.SpecialCells(xlLastCell)).Copy

    ThisWorkbook.Activate
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Workbooks("Book1.xlsx").Close SaveChanges:=False
    Kill ("D:\Users\sb048948\Downloads\Book1.xlsx")
    Application.DisplayAlerts = True

End Sub

Sub copiar_Temp_para_Pedidos()


    Sheets("Temp").Activate
    Range(Range("A2"), ActiveCell.SpecialCells(xlLastCell)).Copy
    Sheets("Pedidos Emitidos JDE").Activate
    Range("E1").End(xlDown).Select
    Selection.Offset(1, -4).Select
    ActiveSheet.Paste
    driver.ExecuteScript ("javascript:onClick=JDEDTAFactory.getInstance('').post('gLS0_1')" )

End Sub

