# -*- coding: utf-8 -*-
"""
Created on Tue Sep 11 08:30:49 2018

@author: marcos.souto
"""

import pandas as pd
import os
import datetime
import Email_Senhas

my_dir = str(r'\\GcRJ01T-SRV03\Privado\Gestão de Risco e Saúde\VALIDAÇÕES\CNU\Valid senhas')
os.chdir(my_dir)


# --- cria data de refecencia para filtrar base de senhas -------------------------#
now = datetime.datetime.now()
try:
    with open(r'C:\Users\marcos.souto\Desktop\Senha\refData.txt', 'r') as arq:
        refData = arq.readlines()
    if now.hour < 12:
        dtREF = datetime.datetime(refData) + datetime.timedelta(days=1)
        with open(r'C:\Users\marcos.souto\Desktop\Senha\refData.txt', 'w') as arq:
            arq.write(str(dtREF))
    else:
        dtREF = datetime.datetime(refData)
except Exception:
    Ref = now - datetime.timedelta(days = 2)
    dtREF = Ref.date()
    with open(r'C:\Users\marcos.souto\Desktop\Senha\refData.txt', 'w') as arq:
        arq.write(str(dtREF))
#-----------------------------------------------------------------------------------

# ---  Abre as bases e carrega para do DataFrame ------------------------------------
df = pd.read_excel('./Senhas_Valid.xlsm', sheet_name = "Geral")
dfCluster = pd.read_excel('./Resumo_Cluster_Valid.xlsx', sheet_name = "Classificados")

df['DT_OCORRENCIA'] = pd.to_datetime(df['DT_OCORRENCIA'])

# --- Filtra os DataFrame --------------------------------------------------------------
dfFitrado = df.loc[lambda df: df.DT_OCORRENCIA > dtREF, :]
dfClusteFiltrado = dfCluster.loc[lambda dfCluster:dfCluster.classificacao != 1, :]

#--- loop para verificar se existe senhas dos beneficiario flegados
lista = []
for row1 in dfCluster.itertuples():
    for row2 in dfFitrado.itertuples():
        if row2.NOME_ASSOCIADO == row1.Nome_Beneficiario:
            lista.append([row2.NUM_ASSOCIADO,
                          row2.NOME_ASSOCIADO,
                          row2.DT_OCORRENCIA,
                          row2.NOME_PRESTADOR,
                          row2.NOME_TRATAMENTO,
                          row2.NOME_PROCEDIMENTO,
                          row1.classificacao])

#--- Verifica se encontrou alguma senha e envia e-mail 
if not lista:
    msg = "<p>E-mail enviado pela gest&atilde;o de senhas Grupo Case.</p><p> N&atilde;o foram encontradas senhas para os benefici&aacute;rios flegados como alerta.</p>"
    Email_Senhas.email_sem_anexo(msg)

else:   
    dfSenhas = pd.DataFrame(lista)
    dfSenhas.columns = ['NUM_ASSOCIADO',
                        'NOME_ASSOCIADO',
                        'DT_OCORRENCIA',
                        'NOME_PRESTADOR',
                        'NOME_TRATAMENTO',
                        'NOME_PROCEDIMENTO',
                        'classificacao']
    dfSenhas.to_excel('./Anexo_Senha.xls')

    msg = "<p>E-mail enviado pela gest&atilde;o de senhas Grupo Case.</p><p> No anexo est&atilde;o as senhas do ultimo dois dias dos benefici&aacute;rios flegados para alerta.</p>"

    Email_Senhas.adiciona_anexo(msg, './Anexo_Senha.xls')


    body = 'Nosso programa acabou de detectar solicitação de homecare. <br><br> Informações abaixo: <br>'
    Email_Senhas.send_email(recipients="marcos.souto@grupocase.com.br",
                            subject="Gestão de Senhas",
                            body=body + './Anexo_Senha.xls',
                            isPlainText=False)
