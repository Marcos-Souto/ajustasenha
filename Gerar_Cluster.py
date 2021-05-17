# -*- coding: utf-8 -*-
"""
Created on Wed Oct  3 11:43:53 2018

@author: marcos.souto
"""
import numpy as np
import pandas as pd
import re
from sklearn.cluster import KMeans
from sklearn.datasets import make_blobs
import random

random.seed(2008)
myDir = str(r'C:\Users\marcos.souto\OneDrive - Dinamica Administração e '
            'Corretagem de Seguros Ltda\Clientes\Valid\AnalisePS\Base.csv')

dados = pd.read_csv(myDir, engine='python', sep=';')

# -- Cria um subset ---
dadosx = dados.iloc[:, [3, 4, 5, 6, 0, 7, 8]]

# -- Renomeio as colunas
dadosx.columns = ['Nome_Beneficiario',
                  'Carteira',
                  'Sexo',
                  'Tipo',
                  'CETIPO',
                  'Qtd_Procedimento',
                  'Sinistro']

# -- substitui virgula por ponto do sinistro.
dadosx['Sinistro'] = pd.to_numeric(dadosx['Sinistro'].apply(
        lambda x: re.sub(",", ".", str(x))))

# --  Cria uma pivot table para transformar a linha do CETIPO em colunas
df = dadosx.pivot_table(index=['Nome_Beneficiario',
                               'Carteira',
                               'Sexo',
                               'Tipo',
                               'Sinistro'], columns='CETIPO', aggfunc='sum')

# -- Grava o pivot Table como DataFrame
dff = pd.DataFrame(df.to_records())

# -- corrige o nome das colunas
dff.columns = [hdr.replace("('Qtd_Procedimento', ", "").replace(")", ""
                           ).replace("'", "") for hdr in dff.columns]

# -- Cria coluna com o custo da internação
dff['Custo_Interna'] = dff['InternaÃ§Ã£o']*dff['Sinistro']

# -- Agrega os valores somando a quantidade de procedinemtos e sinistro
df = dff.groupby(['Carteira']).sum()
dfC = pd.DataFrame(df.to_records())

# -- roda o Kmeans com 5 clusters
y_pred = KMeans(n_clusters=5, random_state=100).fit_predict(dfC.iloc[:, 1:7])

# -- cria nova coluna do DataFrame com a classificação
dfC['Classificado'] = y_pred
# -- Grava o DataFrame como csv
dfC.to_csv(r'C:\Projetos\Python\Arquivos_Uteis\testeKmean.csv', sep=';')
