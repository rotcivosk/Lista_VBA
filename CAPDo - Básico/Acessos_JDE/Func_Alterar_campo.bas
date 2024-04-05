Function alterar_campo(campo As String, dado_inputado As String, tipo_proc_element as string)

    Select case tipo
    case "ID"
        With driver
            .FindElementById(campo, timeout:=10000).Clear
            .ExecuteScript ("document.getElementById('" & campo & "').value = '" & dado_inputado & "'")
            .FindElementById(campo, timeout:=10000).SendKeys Enter
        End With
    case "Name"
        With driver
            .FindElementByName(campo, timeout:=10000).Clear
            .ExecuteScript ("document.getElementByName('" & campo & "').value = '" & dado_inputado & "'")
            .FindElementByName(campo, timeout:=10000).SendKeys Enter
        End With
    case Else
        debug.print "Mandou algo de errado aí no" & campo
    end Select

End Function