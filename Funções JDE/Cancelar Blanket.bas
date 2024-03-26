Sub Cancelar_Blanket()
    
    Dim texto As String
    Call Abrir_Chrome("http://sahdamvpjde009.sa.mds.honda.com:71/jde/E1Menu.maf?jdeLoginAction=LOGOUT&RENDER_MAFLET=E1Menu")
    Call Login_jde("SB048948", "Compras@98")
    
    
    
    '----CATÁLOGO----
    
    With driver
        .FindElementById("drop_fav_menus").Click
        .FindElementByXPath("//div[3]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td/span").Click
        .FindElementByLinkText("Consulta Blanket Order").Click
    End With
    
    Application.Wait (Now + TimeValue("0:00:07"))
    driver.SwitchToFrame 8
    
    For Each cell In Range(Range("AA10"), Range("AA10").End(xlDown))
    
        texto = ""
        With driver
            .FindElementById("C0_13").Clear
            .FindElementById("C0_13").SendKeys cell.Value
            .FindElementById("hc_Find").Click

        End With

        Application.Wait (Now + TimeValue("0:00:01"))
        driver.FindElementById("selectAll0_1", timeout:=10000).Click
        driver.FindElementById("hc_Select", timeout:=10000).Click
        
        Application.Wait (Now + TimeValue("0:00:03"))
        texto = driver.FindElementById("C0_16", timeout:=10000).Value
        

            If texto = "" Then
                texto = driver.FindElementById("C0_231", timeout:=10000).Value
                driver.FindElementById("C0_16", timeout:=10000).Clear
                driver.FindElementById("C0_16", timeout:=10000).SendKeys texto
            End If
            Application.Wait (Now + TimeValue("0:00:01"))
            driver.FindElementById("hc_OK", timeout:=10000).Click
            Application.Wait (Now + TimeValue("0:00:01"))
            driver.FindElementById("hc_OK", timeout:=10000).Click
            Application.Wait (Now + TimeValue("0:00:01"))
            driver.FindElementById("selectAll0_1").Click
            driver.FindElementById("divC0_755", timeout:=10000).Click
            Application.Wait (Now + TimeValue("0:00:01"))
            driver.FindElementByXPath("//body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[7]/div[1]/table/tbody/tr/td/div[6]/table/tbody").Click
            Application.Wait (Now + TimeValue("0:00:01"))
            driver.FindElementById("hc_OK", timeout:=10000).Click
            Application.Wait (Now + TimeValue("0:00:02"))
            driver.FindElementById("hc_OK", timeout:=10000).Click
            Application.Wait (Now + TimeValue("0:00:02"))
            driver.FindElementById("hc_OK", timeout:=10000).Click
            'Application.Wait (Now + TimeValue("0:00:08"))
            cell.Offset(0, 1).Value = "Cancelado"
    Next
    
    Call fechar_Chrome


End Sub
