# Processo seletivo TrackCash

O objetivo deste repositório é armazenar os códigos necessários para o processo seletivo da empresa TrackCash. Os três algoritmos consistem em:
- emailAutomation: extração de anexo de email usando o assunto como critério.
- seleniumExtraction: automação de rotina utilizando a biblioteca Selenium.
- spreadsheetAnalysis: análise de dados utilizando Pandas.

Observações: Para rodar os algoritmos é necessário que seja configurado as variáveis de ambiente nos arquivos assets/.env_email e assets/.env_selenium

## Explicação algoritmos

### emailAutomation

    import email
    import imaplib
    from dotenv import load_dotenv
    from os import environ,path,getcwd
    
   A biblioteca imaplib é utilizada para acessas o conteúdo da caixa de email através do host, a email é usada para converter o formato do e-mail de bytes para base64. A biblioteca dotenv é uma biblioteca para gerenciar variaveis de ambiente, a biblioteca os é utilizada para gerenciar os paths do projeto.

    envPath = path.join(getcwd(),'assets','.env_email')
    
    load_dotenv(envPath)
Variáveis globais usadas no projeto, basicamente carrega as variáveis de ambiente salvas no arquivo assets/.env_email.

    def connectOnEmail():
	    emailHost = environ.get("HOST")
	    mail = imaplib.IMAP4_SSL(emailHost)
	    mail.login(environ.get("EMAIL_USER"),environ.get("EMAIL_PASSWORD"))
    
	    return mail
        
   Função para se conectar no e-mail fornecido nas variáveis de ambiente, usando o host declarado no mesmo.

       def disconnectFromEmail(mail):
	       mail.close()
		   mail.logout()
		   
  Função que desconecta do e-mail passado como parâmetro.

    def getAttacheament(filename,content):
 
	    if filename != None and content != None:
	    
		    savePath = path.join(getcwd(),'assets','extraction',filename)
	    
	    with open(savePath,'wb') as emailAttacheament:
		   emailAttacheament.write(content.get_payload(decode=True))
		   emailAttacheament.close()
  Função para baixar e salvar o arquivo anexo do e-mail selecionado, recebe como parâmetro o nome do arquivo e o seu conteúdo.

    def searchMessage(mail,subject):
    
	    mail.select()
	    
	    typ,data = mail.search(None,'(SUBJECT "{subject}")'.format(subject=subject))
	    
	    if typ == 'OK':
	    
		    typ,data = mail.fetch(data[0],'(RFC822)')
			emaiBody = data[0][1]
			emailFromBytes = email.message_from_bytes(emaiBody)
		    
		    for part in emailFromBytes.walk():
		    
			    if part.get_content_maintype() == "multipart":
				    continue
	    
			    if part.get('Content-Disposition') is None:
				    continue
		    
			    filename = part.get_filename()
			    
			    return filename,part
		    
	    return [None,None]

Função para procurar um e-mail na caixa do usuário logado usando como critério o assunto, essa função retorna o a parte do e-mail que contém o anexo e também o nome do anexo.

    if __name__ == '__main__':
        mail = connectOnEmail()
    	filename,part = searchMessage(mail,"Planilha de Repasse")
        getAttacheament(filename,part)
        disconnectFromEmail(mail)
Função principal, utilizada para controlar toda a aplicação, no primeiro comando ele faz uma instancia de um cliente de e-mail já logado nas credenciais das variáveis de ambiente, logo depois ele chama a função de procurar e-mail utilizando o objeto de e-mail e o parâmetro de busca por assunto, seguindo para o terceiro comando ele baixa o anexo utilizando as variáveis extraídas da função anterior e em sequência desconecta do e-mail.

### spreadsheetAnalysis

    import pandas as pd
    from os import getcwd,path
A biblioteca pandas será utilizada para analisar os dados da planilha baixada no ultimo algoritmo já a biblioteca os continuará com a função de gerenciar os paths do projeto.

    def loadSpreadSheet(spreadsheetName):
    
	    spreadSheetPath = path.join(getcwd(),'assets','extraction',spreadsheetName)
	    
	    if path.exists(spreadSheetPath):
	    
		    spreadsheet = pd.read_excel(spreadSheetPath)
		    spreadsheet = spreadsheet.fillna(0.0)
		    
		    return spreadsheet
Essa função carregará a planilha para uma estrutura de DataFrame, além de tratar os dados para que todos os valores NaN sejam substituídos por 0, evitando possíveis erros nas operações.

    def processInfos(spreadsheet):
    
	    print("Linhas processadas:"+str(len(spreadsheet)))
	    print("Registros Conciliados:"+str(len(spreadsheet[spreadsheet["Conciliação*"] == "Conciliado"])))
	    print("Registros Não Conciliados:"+str(len(spreadsheet[spreadsheet["Conciliação*"] == "Não Conciliado"])))
	    print("Registros Retirados:"+str(len(spreadsheet[spreadsheet["Conciliação*"] == "Retirada"])))
	    print("Registros Movimentados:"+str(len(spreadsheet[spreadsheet["Conciliação*"] == "Movimentação"])))
	   
	    for column in ["Comissão ML por parcela","Valor bruto da parcela"]:
	    
		    if spreadsheet[column].dtypes == 'float64' or spreadsheet[column].dtype == 'int64':
			    sum = str(spreadsheet[column].sum())
			    print("Soma da coluna {column}:{sum}".format(column=column,sum=sum))
Essa função irá exibir alguns dados a respeito da planilha, como a soma de colunas numéricas utilizadas na visualização além de quantidade de dados.

    def compareMultiplicationWithInput(multiplicand,multiplicator,inputProduct):
    
	    if multiplicand * multiplicator == inputProduct:
		    return "Conciliado"
	    
	    return "Não Conciliado"
Função que será utilizado pela funções que serão mostradas a seguir para definir uma transação como "Conciliado" ou "Não Conciliado", de acordo com as regras de negócio passadas.

    def creditCardAndBillet(grossInstallmentAmount,commission,mlComissionPerInstallment):
    
	     return compareMultiplicationWithInput(grossInstallmentAmount,commission,mlComissionPerInstallment)
    
    def reversal(grossValueOrder,comission,liquidValueInstallament):
    
	    return compareMultiplicationWithInput(grossValueOrder,comission,liquidValueInstallament)
Em suma as duas funções executam a mesma operação, deixei elas separadas somente para facilitar a leitura e entendimento do código. A primeira função trata o caso específico de pagamento com cartão de crédito ou boleto e a segunda no caso de estorno.

    def transference(antecipationValue):
    
	    if antecipationValue>0:
		    return "Retirada"
	    
	    return "Movimentação"
Função para tratar o caso específico de transferência, classificando a conciliação como "Retirada" ou "Movimentada" de acordo com a regra de negócio.

    def specificCases(row):
    
	    conciliation = None
	    
	    if row["Método de pagamento"].lower() == "cartão de crédito" or row["Método de pagamento"].lower() == "boleto":
			conciliation = creditCardAndBillet(row["Valor bruto da parcela"],row["% Comissão"],row["Comissão ML por parcela"])
	    
	    elif row["Método de pagamento"].lower() == "estorno":
		    conciliation = reversal(row["Valor bruto do pedido"],row["% Comissão"],row["Valor líquido da parcela"])
		    
	    elif row["Método de pagamento"].lower() == "transferência":
		    row["Comissão ML por parcela "] = row["Valor da antecipação"]
		    row["Valor bruto da parcela"] = row["Valor líquido da parcela"]
		    
		    conciliation = transference(row["Valor da antecipação"])
	    
	    if conciliation != None:
		    row["Conciliação*"] = conciliation
	    
	    return row
Função que trata dos casos específicos da linha que é usada como parâmetro, de acordo com o "Método de Pagamento" a conciliação será tratada com uma função diferente, a função retorna a linha de entrada com o acréscimo da coluna de conciliação.

    def processRows(spreadsheet):
    
	    processedDataFrame = pd.DataFrame()
	    for index,row in spreadsheet.iterrows():
	    
		    treatedRow = specificCases(row)
		    print(treatedRow[["Data da transação","ID do pedido Seller","Método de pagamento","Comissão ML por parcela","Valor bruto da parcela","% Comissão","Conciliação*"]])
		    print('------------')
		    
		    processedDataFrame = processedDataFrame.append(treatedRow)
	    
	    return processedDataFrame
Função para acessar cada registro da base de dados, chamar a função que define a conciliação e em seguida imprimir na tela o registro somente com as colunas exigidas na regra de negócio, além disso um novo DataFrame com a coluna "Conciliação*" é criado para depois ser analisado.

    if __name__ == '__main__':
	    spreadsheet = loadSpreadSheet('planilha_de_repasse.xlsx')
	    newSpreadSheet = processRows(spreadsheet)
	    processInfos(newSpreadSheet)
Função principal, no primeiro comando irá carregar os dados da planilha, em sequência será rodado o comando para processar as linhas e retornar um novo DataFrame com a coluna "Conciliação*", por ultimo será rodado o comando para exibir informações a respeito dos dados analisados.

### seleniumExtraction

    from dotenv import load_dotenv
	from os import environ,path,getcwd
    from selenium.webdriver import Chrome
    import time
A biblioteca dotenv será usada para gerenciar as variáveis de ambiente, a biblioteca os será usado para acessas as variáveis de ambiente, a biblioteca selenium será o core do algoritmo controlando toda a automação, a biblioteca time será usado para controlar o tempo entre execução de comandos.

    envPath = path.join(getcwd(),'assets','.env_selenium')
	load_dotenv(envPath)
    webdriverPath = environ.get('WEBDRIVER')
    webdriver = Chrome(webdriverPath)
As variáveis globais basicamente irão carregar as variáveis de ambiente e fazer a instância do webdriver do Chrome.

    def acessPage(path):
	    webdriver.get(path)
Função para fazer acesso a página passada como parâmetro, apesar de parecer desnecessário é importante para manter o código mais legível.

    def setForms(formCriteria):
	     for key in formCriteria:
		   webdriver.find_element_by_id(key).send_keys(formCriteria[key])
Função para preencher formulários por id, recebe como parâmetro um dicionário com a seguinte estrutura: `{ID_DO_CAMPO:VALOR_A_SER_ENVIADO,...}`

    def arrayCheckClick(array,criteria):
    
        for item in array:
	        try:
		        if item.text == criteria:
			       time.sleep(1)
			        item.click()
        
	        except:
		        continue
Função para verificar diversos botões em uma lista e clicar nele caso seu valor de texto seja satisfeito pelo parâmetro "criteria".

    def removeDuplicateDays(array):
    
	    monthWeeks = False
	    previousValue = 0
	    outputArray = []
	    
	    for item in array:
		    
		    if int(item.text) < previousValue and monthWeeks:
			    break
		    
		    if int(item.text) > 1 and monthWeeks == False:
			    pass
		    
		    else:
			   previousValue = int(item.text)
			   monthWeeks = True
			   outputArray.append(item)
	    
	    return outputArray
Essa função possui uma tarefa bem específica que é remover valores duplicados nos botões para selecionar os dias do mês no componente de calendário do portal que é acessado, evitando inconsistências, principalmente quando for ser selecionado valores como "31" e "1".

    def setCalendar(calendarCriteria):
    
	    iFrameLocation = webdriver.find_element_by_id("main").find_element_by_class_name("container-fluid")
	    time.sleep(3)
		webdriver.switch_to.frame(iFrameLocation.find_element_by_tag_name("iframe"))
	    
	    calendarInput = webdriver.find_element_by_class_name("emulated-calendar__component")
	    calendarInput.click()
	    webdriver.find_element_by_class_name("react-calendar__navigation__label").click()
	    webdriver.find_element_by_class_name("react-calendar__navigation__label").click()
	    
	    yearButtons = webdriver.find_elements_by_class_name("react-calendar__decade-view__years__year")
	    arrayCheckClick(yearButtons,calendarCriteria["year"])
	    
	    monthButtons = webdriver.find_elements_by_class_name("react-calendar__year-view__months__month")
	    arrayCheckClick(monthButtons,calendarCriteria["month"])
	    
	    daysButtons = webdriver.find_elements_by_class_name("react-calendar__month-view__days__day")
		daysButtons = removeDuplicateDays(daysButtons)
	    arrayCheckClick(daysButtons,calendarCriteria["begin"])
	    arrayCheckClick(daysButtons,calendarCriteria["end"])
	    
	    exportExcel = webdriver.find_element_by_class_name("MuiButton-outlined")
	    exportExcel.click()
	    
	    time.sleep(2)
	    closeModal = webdriver.find_element_by_class_name("jss79").find_element_by_tag_name("button")
	    closeModal.click()
Essa função será encarregada de manipular o componente de calendário do portal, selecionando o ano,mês e dias em que será realizado a consulta e em sequência exportando os dados atráves do botão de exportação.Essa função recebe um dicíonario como parâmetro com os valores que irá usar no componente de calendário, o dicíonario possui a seguinte estrutura:`{"year":ano,"month":mês,"begin":primeiro dia,"end":ultimo dia}`

    def clickButton(tag_id):
	    webdriver.find_element_by_id(tag_id).click()
Função para clicar em botão, orientado por id.

    if __name__ == '__main__':
    
	    acessPage(environ.get("LOGIN_URL"))
	    
	    setForms({
		    "username":environ.get("EMAIL"),
		    "password":environ.get("PASSWORD")
	    })
	    
	    clickButton("kc-login")
	    acessPage(environ.get("ORDER_URL"))
	    setCalendar({
		    "year":"2018",
		    "month":"janeiro",
		    "begin":"1",
		    "end":"31"
		})
		webdriver.close()
A função principal executará primeiro o serviço para acessar a pagina de login, em sequência irá preencher o formulário da pagina de acordo com os parâmetros passados e clicará no botão de login e acessará a pagina de pedidos, em sequência executará a função para manipular o componente de calendário e exportar os dados, por ultimo irá fechar o webdriver.
