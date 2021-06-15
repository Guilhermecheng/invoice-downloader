import pandas as pd
import math

path = r"D:\downpdf\DOWNLOAD_SAP.xlsx"
cols = [0,1,3,4]
df = pd.read_excel(path, engine='openpyxl', usecols=cols)

# print(df)
ordem_de_servico = []
nota_1_lista = []
nota_2_lista = []
loja_lista = []
for row in range(len(df)):
    ordem = str(df['OS_Numero'][row])
    loja = df['Empresa_Nome'][row]
    nota_1 = df['OS1'][row]
    nota_2 = df['OS2'][row]
    if math.isnan(nota_1) == False:
        nota_1 = str(int(nota_1))
    else:
        nota_1 = "none"

    if math.isnan(nota_2) == False:
        nota_2 = str(int(nota_2))
    else:
        nota_2 = "none"

    ordem_de_servico.append(ordem)
    nota_1_lista.append(nota_1)
    nota_2_lista.append(nota_2)
    loja_lista.append(loja)

    
