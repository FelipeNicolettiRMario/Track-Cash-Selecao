import pandas as pd
from os import getcwd,path

def loadSpreadSheet(spreadsheetName):
    spreadSheetPath = path.join(getcwd(),'assets','extraction',spreadsheetName)

    if path.exists(spreadSheetPath):
       spreadsheet = pd.read_excel(spreadSheetPath)
       spreadsheet = spreadsheet.fillna(0.0)

       return spreadsheet

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


def compareMultiplicationWithInput(multiplicand,multiplicator,inputProduct):

    if multiplicand * multiplicator == inputProduct:
        return "Conciliado"

    return "Não Conciliado"

def creditCardAndBillet(grossInstallmentAmount,commission,mlComissionPerInstallment):

    return compareMultiplicationWithInput(grossInstallmentAmount,commission,mlComissionPerInstallment)

def reversal(grossValueOrder,comission,liquidValueInstallament):

    return compareMultiplicationWithInput(grossValueOrder,comission,liquidValueInstallament)

def transference(antecipationValue):
    if antecipationValue>0:
        return "Retirada"

    return "Movimentação"

def specificCases(row):

    conciliation = None
    if row["Método de pagamento"].lower() == "cartão de crédito" or  row["Método de pagamento"].lower() == "boleto":
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

def processRows(spreadsheet):

    processedDataFrame = pd.DataFrame()
    for index,row in spreadsheet.iterrows():

        treatedRow = specificCases(row)
        print(treatedRow[["Data da transação","ID do pedido Seller","Método de pagamento","Comissão ML por parcela","Valor bruto da parcela","% Comissão","Conciliação*"]])
        print('------------')
        processedDataFrame = processedDataFrame.append(treatedRow)

    return processedDataFrame

if __name__ == '__main__':
    spreadsheet = loadSpreadSheet('planilha_de_repasse.xlsx')
    newSpreadSheet = processRows(spreadsheet)
    processInfos(newSpreadSheet)