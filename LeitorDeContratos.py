from PyPDF2 import PdfReader
import re
import pandas
import os
import datetime

DfDataColumns = {
    "CONTRATO": [],
    "NOME CONTRATANTE":[],
    "ID":[],
    "CPF":[],
    "ENDEREÇO CONTRATANTE":[],
    "EMAIL":[],
    "TELEFONE":[],
    "DATA EVENTO":[],
    "VALOR CONTRATO":[],
    "FORMA DE PAGAMENTO":[]
}

DfContratos = pandas.DataFrame(DfDataColumns)

directory = '.\\Contratos'

for filename in os.listdir(directory):
    file = os.path.join(directory, filename)

    print("Extraindo dados do Contrato: " + filename)

    # Extraindo texto dos contratos
    reader = PdfReader(file)
    pages = reader.pages
    text = ""
    for page in pages:
        text += page.extract_text().replace("\n", "")
    textoSemHeader = text.split("CONTRATADO, e do")[1]

    # REGEX para extrair os dados
    nome = re.findall('(?s)(?<=outro lado, ).*?(?=, )', textoSemHeader)[0]
    id = str(re.findall('(?s)(?<=, ID ).*?(?=, CPF)', textoSemHeader)[0]).replace("-","").replace(".","")
    cpf = str(re.search('(?s)(?<=, CPF ).*?(?=( Endereço)|( Nº de contato)|( email))', textoSemHeader)[0]).replace("-","").replace(".","")

    # Endereço
    endereco = ""
    try:
        endereco = re.search('(?s)(?<=Endereço ).*?(?=( email)|( Nº de contato))', textoSemHeader)[0]
    except:
        endereco = ""

    # Email
    email = ""
    try:
        email = re.search('(?s)(?<=email ).*?(?=(, Nº de contato )|(, cel)|(, doravante )|( doravante ))', textoSemHeader)[0]
    except:
        email = ""

    # Telefone
    telefone = ""
    try:
        telefone = str(re.search('(?s)(?=(Nº de contato )|(cel )).*?(?=( doravante )|(, doravante )|( email))', textoSemHeader)[0]).replace("cel ","").replace("Nº de contato ","").replace("(","").replace(")","").replace("-","").replace(" ","")
    except:
        telefone = ""

    # DataEvento
    dataEvento = ""
    try:
        dataEvento = re.findall('(?s)(?<=que acontecerá no dia ).*?(?=[, ])', textoSemHeader)[0]
    except:
        dataEvento = ""
    
    # Demais dados
    valorContrato = re.findall('(?s)(?<=valor total de ).*?(?=, sendo)', textoSemHeader)[0]
    formaDePagamento = re.findall('(?s)(?<=, sendo feito o pagamento ).*?(?=4a Cláusula)', textoSemHeader)[0]

    DfContratos.loc[len(DfContratos.index)] = [filename, nome, id, cpf, endereco, email, telefone, dataEvento, valorContrato, formaDePagamento]

RelatorioName = "Relatório de Contratos {0}.xlsx".format(str(datetime.datetime.now()).replace("-","_").replace(":","").replace(".",""))
PathToExport = '.\\Relatório\\' + RelatorioName
DfContratos.to_excel(PathToExport, index = False)