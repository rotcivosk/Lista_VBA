Set arquivo_sistema = CreateObject("Scripting.FileSystemObject")
dim caminho_arquivos
set caminho_arquivos = "D:\Users\sb048948\OneDrive - Honda\Documentos\Baixar Arquivos\"
Set tempFile = arquivo_sistema.OpenTextFile(caminho_arquivos & "Temp\pedido.txt")
dim pedido
pedido = tempfile.Readline() ' Recebe o valor do pedido como argumento
tempFile.Close

Set WShell = CreateObject("WScript.Shell")
n = 1
Do
    achou_tela = WShell.AppActivate("Imprimir")
    n = n + 1
    WScript.Sleep 100
Loop Until achou_tela Or n > 100

With WShell
    .AppActivate "Imprimir"
    .SendKeys "%{N}"
    .SendKeys "%{DOWN}"
    .SendKeys "{UP}"
    .SendKeys "{UP}"
    .SendKeys "{UP}"
    .SendKeys "{UP}"
    .SendKeys "{DOWN}"
    .SendKeys "{DOWN}"
    .SendKeys "{DOWN}"
    .SendKeys "{ENTER}"
    .SendKeys "{ENTER}"
End With

Set WShell = Nothing
Set WShell = CreateObject("WScript.Shell")
n = 1
Do
    achou_tela = WShell.AppActivate("Salvar Saída de Impressão como")
    n = n + 1
    WScript.Sleep 10
Loop Until achou_tela Or n > 100
WScript.Sleep 100

With WShell
    .AppActivate "Salvar Saída de Impressão como"
    .SendKeys "+{TAB}"
    .SendKeys "{TAB}"
    .SendKeys pedido
    .SendKeys "{F4}"
    .SendKeys "{DOWN}"
    .SendKeys "{DELETE}"
    .SendKeys caminho_arquivos
    .SendKeys "{ENTER}"
End With
WScript.sleep 1000

With WShell
    .AppActivate "Salvar Saída de Impressão como"
    .SendKeys "{ENTER}"
end with

With WShell
    .AppActivate "Salvar Saída de Impressão como"
    .SendKeys "{ENTER}"
end with

With WShell
    .AppActivate "Salvar Saída de Impressão como"
    .SendKeys "{ENTER}"
end with
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(caminho_arquivos & "temp\flag.txt", True)
objFile.Write "done"
objFile.Close