Sub Cancelar_Blanket()
    
    dim user as string, senha as string
    Dim texto As String
    dim titulo_pag as string
    Call Abrir_Chrome("http://sahdamvpjde009.sa.mds.honda.com:71/jde/E1Menu.maf?jdeLoginAction=LOGOUT&RENDER_MAFLET=E1Menu")
    Call Login_jde(user, senha)
    
    
    
    '----CATÁLOGO----
    
    With driver
        .FindElementById("drop_fav_menus").Click
        .FindElementByXPath("//div[3]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td/span").Click
        .FindElementByLinkText("Consulta Blanket Order").Click
    End With
    
    Application.Wait (Now + TimeValue("0:00:07"))
    driver.SwitchToFrame 8
    
    For Each cell In Range(Range("AA10"), Range("AA10").End(xlDown))
    
        ' Check para evitar Duplicidades
        If cell.Offset(0,1) = "" Then

            ' Gerenciador de Erros
            if esperar_carregar(" Consulta Blanket Order - Acesso a Detalhes de Pedidos") Then GoTo Deu_erro
            texto = ""
            numero_ped = cell.value

            ' Adicionar valor
            With driver
                .FindElementById("C0_13").Clear
                ' Alterado de Sendkeys para execute script para maior confiabilidade
                .ExecuteScript ("document.getElementById('C0_13').value = '" & numero_ped & "'")
                .FindElementById("hc_Find").Click

            End With

            Application.Wait (Now + TimeValue("0:00:02"))

            ' Checar fornecedor antes de abrir o processo
            fornecedor = driver.FindElementByXPath("/html/body/form[3]/div/table/tbody/tr/td/div/span[9]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td[11]/div").Text
            forn_planilha = Range("AB7").Text

            If fornecedor = forn_planilha Then

                ' Seleciona a primeria linha (Melhor que selecionar todas)
                driver.FindElementByXPath("//html/body/form[3]/div/table/tbody/tr/td/div/span[9]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td/div/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td[1]/div/input").Click
                Application.Wait (Now + TimeValue("0:00:01"))
                driver.FindElementById("hc_Select", timeout:=10000).Click               
                Application.Wait (Now + TimeValue("0:00:03"))

                ' Carregar a página
                If esperar_carregar(" Consulta Blanket Order - Cabeçalho do Pedido") Then GoTo Deu_erro                

                ' Alterar valores de datas iniciais
                texto = driver.FindElementById("C0_231", timeout:=10000).Value
                driver.FindElementById("C0_16", timeout:=10000).Clear
                driver.ExecuteScript ("document.getElementById('C0_16').value = '" & texto & "'")
                Application.Wait (Now + TimeValue("0:00:01"))
                
                ' Algumas vezes é necessário selecionar duas vezes
                If driver.FindElementById("jdeFormTitle0", timeout:=10000).Text = " Consulta Blanket Order - Cabeçalho do Pedido" Then
                    driver.FindElementById("hc_OK", timeout:=10000).Click
                End If

                ' Checar se a página cerregou, selecionar todas as linhas e cancelar
                If esperar_carregar(" Consulta Blanket Order - Detalhes do Pedido") Then GoTo Deu_erro                
                
                driver.FindElementById("selectAll0_1", timeout:=10000).Click
                driver.FindElementById("divC0_755", timeout:=10000).Click
                driver.FindElementByXPath("//body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[7]/div[1]/table/tbody/tr/td/div[6]/table/tbody", timeout:=10000).Click
                Application.Wait (Now + TimeValue("0:00:01"))
                driver.FindElementById("hc_OK", timeout:=10000).Click    
                
                ' às vezes pressisa ser pressionado duas vezes
                If driver.FindElementById("jdeFormTitle0", timeout:=10000).Text = " Consulta Blanket Order - Detalhes do Pedido" Then
                    driver.FindElementById("hc_OK", timeout:=10000).Click
                End If

                ' Primeira Tela e OK
                If esperar_carregar(" Inf. Adicionais de Detalhes de Pedidos de Compras - Brasil") Then GoTo Deu_erro                
                driver.FindElementById("hc_OK", timeout:=10000).Click

                ' Segunda Tela e OK
                If esperar_carregar(" Consulta Blanket Order - Objeto de Mídia") Then GoTo Deu_erro                
                driver.FindElementById("hc_OK", timeout:=10000).Click

                ' Confirmar que foi cancelado e seguir para o próximo
                If esperar_carregar(" Consulta Blanket Order - Acesso a Detalhes de Pedidos") Then GoTo Deu_erro    
                cell.Offset(0, 1).Value = "Cancelado"

            else
                cell.Offset(0,1).value = "Fornecedor Incorreto"
            end if
        end if  
        erro = false
        if erro Then
            Deu_erro:
            cell.Offset(0, 1).Value = "Falha no cancelamento"
            exit Sub ' Decorrente do sistema, necessário reiniciar o sistema
        end if  
    Next
    
    Call fechar_Chrome

End Sub