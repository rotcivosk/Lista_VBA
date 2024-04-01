
0 - Limpar
	Limpar tabela ME5A
	Limpar tabela ME2N
	Criar / Limpar tabela TEMP

1 - ME5A
	Filtro Inicial é padrão no user SB048948	
	Output são as requisições em aberto no momento
	Informações Esperadas:
	- N° da Requisição e Linha
	- Cód. Class Contábil
	- Código do Material e Descrição do Material
	- Quantidade, Unidade de MEdida e Valor Líquido
	- Grupo de Mercadoria
	- Filial
	- Split Local e Nacional
	- N° de RegInfo
	Reorganizar colunas conforme padrão de exportação da ME2N


2 - ME2N
	Filtro Inicial é padrão no user SB048948
	Filtro adicional é a data Inicial e data de hoje do Relatório
	Selecionar as divisões de remessa ao exportar
	Informações Esperadas:
	- N° do Pedido e Linha
	- N° da Requisição e Linha
	- N° do RegInfo
	- Cód. Class Contábil
	- Código do Fornecedor e Descrição do Fornecedor
	- Código do Material e Descrição do Material
	- Data do Pedido Emitido
	- Grupo de Mercadoria
	- Filial
	- Quantidade, Unidade de Medida e Valor Líquido
	Converter formatação da data de emissão do pedido de 01.01.2015 para 01/01/2015
	Mesclar com a planilha ME5A
	Copiar a coluna de Requisições e colar na planilha TEMP
	

3 - TEMP
	- Remover duplicatas da planilha TEMP
	- Mudar a formatação dos textos para 10 zeros (0000000000)
	Copiar as requisições formatadas para utilizar como filtro na CDPOS
	
4 - CDPOS
	Filtros Iniciais são gravados no user sb048948
	Filtros Adicionais são as requisições da TEMP
	Informação esperada:
	- N° de Alteração do Documento
	Copiar os Números de alteração do documento e exportar como filtro na CDHDR	

5 - CDHDR
	Filtros Iniciais são gravados no user SB048948
	Filtros adicionais são os Números de alteração
	Informação esperada:
	- N° da Requisição
	- Data de Aprovação	
	Converter formatação da data de aprovação da requisição de 01.01.2015 para 01/01/2015
	Remover outras colunas além das presentes na lista informação esperada

6 - ME2N (Com ME5A mesclada)
	Copiar os Números de Pedido e exportar como filtro na EKKO

7 - SE16N - EKKO
	Acessar a EKKO
	Filtros Iniciais devem estar vazios
	N° Máximo de ocorrências deve estar vazio
	Filtros adicionais deve ser os Números de pedido da ME2N
	Informação esperada:
	- N° do Pedido
	- Cód. Condições Documento(Cond. Doc)
	- Usuário requisitante
	Copiar os Números de Cond. Doc e exportar como filtro para a KONV

8 - SE16N - KONV
	Acessar a KONV
	Filtros iniciais devem estar vazios
	Filtros adicionais deve ser o Cond. Doc da EKKO
	Filtros adicionais deve ser o tipo de documento ZPBX e PB00
	Informação esperada:
	- Cod. Condições Documento(Cond. Doc)
	- Linha do Pedido
	- Valor Bruto (Montante)
	- Tipo de Item
	Criar coluna mesclando Cond. Doc, Linha e Tipo de Item
	Copiar os Númesb017074ros de Reg Cond para a planilha Temp

9 - TEMP
	Remover duplicatas
	Mudar a formatação dos textos para 10 zeros (0000000000)
	Copiar os Números de Reg Cond como filtros da KONP

10 - SE16N KONP
	Acessar a KONP
	Filtros iniciais devem estar vazios
	Filtros adicionais deve ser os números de Reg Cond
	Informação Esperada:
	- N° Reg Cond
	- Tipo de Item
	- Valor total (Montante)
	Criar coluna mesclando Reg Cond e Tipo C.


11 - ME2N (Com ME5A mesclada)
	Copiar os Números de Material e colar na planilha TEMP

12 - Temp
	Remover duplicatas dos códigos de material exportados
	Exportar os códigos de material como filtros para a EORD

13 - SE16N - EORD
	Acessar a EORD
	Filtros Iniciais devem estar vazios
	Filtros adicionais devem ser os centros - 0212, 0304 e 0232
	Filtros adicionais devem ser a lista de material da TEMP
	Informação esperada:
	- Código de Material
	- Código de Fornecedor
	Exportar os códigos de material como filtros para a EINA

14 - SE16N - EINA
	Acessar a EINA
	Filtros iniciais devem estar vazios
	Filtros adicionais devem ser a lista de códigos de material da EORD
	Informação esperada:
	- Código de RegInfo
	- Código de Material
	- Status de Cabeçalho (Cancelado ou Não)
	Exportar os códigos de Reginfo como filtros para a EINE

15 - SE16N - EINE
	Acessar a EINE
	Filtros iniciais devem estar vazios
	Filtros adicionais devem ser os centros - 0212, 0304 e 0232
	Filtros adicionais devem ser a lista de Reginfo da EINA
	Informação esperada:
	- Código de Reginfo
	- Centro
	- Status da Linha
	Remover coluna além das informações esperadas
	Remover espaços com %20 " "
	Puxar código do material da EINA com base no N° do RegInfo
	Puxar código do fornecedor da EINA com base no N° do RegInfo
	Puxar status do cabeçalho da EINA com base no N° do RegInfo
	Criar coluna com base no status. "RegInfo Cancelado", "RegInfo OK" e organizar os itens "OK" como primeiros
	Criar coluna consolidando N° do material e centro para ser utilizada como link para procv

16 - Alterações na ME2N
	Criar coluna chamada "Sistema" e classificar todos como SAP
	Dividir a coluna de fornecedores em Código e Descrição
	Remover espaços " " (%20) dos campos de Cód Material e Cód Fornedeor
	Puxar da EINE o status do RegInfo com base no N° de material e Centro.
	Caso status do RegInfo seja ok ou cancelado, puxar da EINE o código do fornecedor com base no N° de material e centro.
	Puxar "Cond. Doc" da EKKO com base no N° de Pedido. Itens da ME5A deverão ser ignorados
	Puxar "Valor Bruto Unitário" da KONV com base no Cond. Doc. e no ZPBX
	Puxar "Reg Cond" da KONV com base no Cond. Doc. e PB00
	Puxar "Valor Inicial Unitário" da KONP com base no Reg Cond. e no ZPBI
	Criar coluna de "Valor Bruto Total" e "Valor Inicial Total" com base nas colunas unitárias multiplicadas pela quantidade
	Criar coluna de "Mérito" com base na diferença entre valores
	Criar coluna que Identifica qual é a categoria de cada item com base em na planilha de carteiras

17 - AS400 - Req_aprov
	Exportar em suplementos o resultado da query reqaprov
	Valores esperados:
	- N° do Pedido
	- Data de Aprovação
	Converter a formatação de 20101997 para 20/10/1997
	Organizar do mais recente para o mais antigo
	Remover Duplicatas
	Mesclar na planilha CDHDR

18 - AS400 - Requisições
	Rodar a query relatojoji com filtros:
	- Tip_pedido : OR, OY, OQ
	- Pro_status : 120, 110
	Exportar em suplementos o resultado da query relatojoji

19 - JDE - Pedidos
	Rodar o relatório na tela de follow de requisições com filtros:
	- Tipo de Pedido: OP, OS, OM, OL
	- Data Inicial: Data Inicial do relatório
	- Data Final: Data em que o relatório está sendo processado
	- Filial: 05001, 10001, 05998, 10998
	Exportar para planilha "Pedidos Emitidos JDE"	
	Filtrar itens base catálogo com base na ausência de N° de Requisição
	Remover linhas de pedido que não tenham nem N° de requisição nem N° de cotação

20 - AS400 - Programado
	Rodar a query relatojoji com filtros:
	- Tipo de pedido : XM, XS, O0, O7, O9, QD
	- Codigo de Item: Diferente de: 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, A, B, D, Y
	- Data de emissão: Base data de relatório
	- Prostatus: Diferente de 980

21 - JDE - Performance
	Utilizar a planilha do último relatório, sem apagar o histórico
	Adicionar na linha abaixo
	Processar a transação Perf sum Negócios	

99 - Para fazer na planilha final
	Puxar data de aprovação da planilha CDHDR com base no N° de Requisição
	Puxar o prazo de cada item com base nos dias úteis entre a aprovação da requisição e a emissão do pedido ou a data do dia final caso a requisição esteja em aberto
	Criar coluna que checa ("Fora do Prazo"/"Dentro do Prazo") com base no lead time versus o prazo de cada item
	Remover N° do Pedido, linha do pedido, data de Emissão do pedido para pedidos emitidos após a data de corte com requisições aprovadas antes da data de corte.
Criar coluna que checa ("Fechado"/"Em Aberto") com base na existência ou não de N° de pedido.
	Criar coluna que checa ("Item Novo"/"Mês Anterior") com base na data de aprovação ser anterior ou posterior a data de corte
