# -*- coding: utf-8 -*-
"""
Created on Mon Sep  3 14:24:18 2018

@author: marcos.souto
"""

import pandas as pd
import os
import glob
import shutil
from log import logger
from Gera_Arq_CSV_Geral import LerCVSGeral as LG
from Gerar_Email import Ver_Senhas as VerSenha

# --------------------- Seta o diretorio Energisa --------------------------- #
my_dir = str(r'\\GcRJ01T-SRV03\Privado\Gestão de Risco e Saúde\VALIDAÇÕES\CNU'\
             '\Energisa_Senhas\Relatorio')

os.chdir(my_dir)
# -- Verifica se o diretorio esta vazio
tx = os.path.exists("*.xls")
if tx == False:
    logger.info("Não existe arquivos da Energisa para serem gravado")

for files in glob.glob('*.xls'):
    try:
        # -- le o arquivo .htm e carrega para um DataFrame
        tabela = pd.read_html(files)
        frm = pd.DataFrame(tabela[0])
        frm.to_csv("./Corrigido/" + files + ".csv", sep=',', index=False,
                   encoding='utf8')
        shutil.move(files, './Original') # move o arquivo para pasta original
        # -- Grava o log para acompanhamento
        logger.info("Arquivo " + files + " gravado com sucesso") 
    except:
        # -- Grava o log para acompanhamento
        logger.error("Erro ao gravar o arquivo " + files)

# -------------------- Seta o diretorio Valid ------------------------------ #
my_dir = str(r'\\GcRJ01T-SRV03\Privado\Gestão de Risco e Saúde\VALIDAÇÕES\CNU'\
             '\VALID senhas\Relatorios')
os.chdir(my_dir)
# -- Verifica se o diretorio esta vazio
tx = os.path.exists("*.xls")
if tx == False:
    logger.info("Não existe arquivos da Valid para serem gravados")

for files in glob.glob('*.xls'):
    try:
        # -- le o arquivo .htm e carrega para um DataFrame
        tabela = pd.read_html(files)
        frm = pd.DataFrame(tabela[0])
        frm.to_csv("./Corrigido/" + files + ".csv", sep=';', index=False,
                   encoding='utf8')
        shutil.move(files, './Original')
        # -- Grava o log para acompanhamento
        logger.info("Arquivo " + files + " gravado com sucesso")
    except:
        # -- Grava o log para acompanhamento
        logger.error("Erro ao gravar o arquivo " + files)

# -- Seta o diretorio onde estão os arquivo csv corrigido
my_dir = str(r'\\GcRJ01T-SRV03\Privado\Gestão de Risco e Saúde\VALIDAÇÕES\CNU'\
             '\VALID senhas\Relatorios\Corrigido')
# -- Chama a função que todos os CSV para um unico arquivo
LG(my_dir)
# -- Seta o diretorio
my_dir2 = str(r'\\GcRJ01T-SRV03\Privado\Gestão de Risco e '\
              'Saúde\VALIDAÇÕES\CNU\VALID senhas')

# --chama a função que verifica existe  benficiario ou procedimentos de alerta
VerSenha(strDiretorio=my_dir2,
         tbSenhas='./Relatorios/Corrigido/Geral/CSV_Geral.csv',
         tbCluster='./Resumo_Cluster_Valid.xlsx',
         SheetCluster='Classificados',
         strTitulo='Gestão de Senhas Valid')
