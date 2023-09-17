import xmltodict
import os
import pandas as pd
from tkinter.filedialog import askdirectory, asksaveasfilename


def coleta_infos(nome_arquivo, conteudo):
    with open(os.path.join(pasta, nome_arquivo), "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)
        if "NFe" in dic_arquivo:
            infos_nf = dic_arquivo["NFe"]['infNFe']
        else:
            infos_nf = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
        numero_nota = infos_nf["@Id"]
        empresa_emissora = infos_nf["emit"]["xNome"]
        nome_cliente = infos_nf["dest"]["xNome"]
        endereco = infos_nf["dest"]["enderDest"]
        if "vol" in infos_nf["transp"]:
            peso = infos_nf["transp"]["vol"]["pesoB"]
        else:
            peso = "Sem carga declarada!"

        conteudo.append([numero_nota, empresa_emissora, nome_cliente, endereco, peso])

pasta = askdirectory(title="Selecione a pasta das notas.")
lista_arquivos = os.listdir(pasta)

colunas = ["Número da NFe", "Empresa Emissora", "Nome do Cliente", "Endereço de Entrega", "Peso"]
conteudo = []

for arquivo in lista_arquivos:
    coleta_infos(arquivo, conteudo)

tabela = pd.DataFrame(columns=colunas, data=conteudo)

# Pergunta onde salvar o arquivo e com qual nome
nome_arquivo = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Planilhas Excel", "*.xlsx")], title="Salvar como")

if nome_arquivo:
    tabela.to_excel(nome_arquivo, index=False)
