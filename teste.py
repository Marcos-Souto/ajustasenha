# -*- coding: utf-8 -*-
"""
Created on Wed Sep 26 15:28:42 2018

@author: marcos.souto
"""

import pandas as pd
import os
import glob
from datetime import date
from log import logger

my_dir = str(r'\\GcRJ01T-SRV03\Privado\Gestão de Risco e Saúde\VALIDAÇÕES\CNU'\
             '\VALID senhas\Relatorios\Corrigido')
os.chdir(my_dir)
# -- Abre o arquivo com todo o historico para ser completado
frame = pd.read_csv('./Geral/CSV_Geral.csv', sep=',', encoding='utf8')

    # -- Verifica se o diretorio esta vazio
tx = os.path.exists("*.csv")
if tx == False:
    #logger.info("Não existe arquivos da InimedRio para serem gravados")
    print("Não existe arquivos da InimedRio para serem gravados")

    # -- Realiza loop no diretorio procurando arquivo htm --#
for files in glob.glob('*.csv'):
    # -- le o arquivo  e carrega para um DataFrame
    tabela = pd.read_csv(files)
    frm = pd.DataFrame(tabela)

    # -- Salva o DataFrame como CSV e depois abre ---
    frm.to_csv('./Geral/ArquivosGeralx.csv', sep=';', index=False)
    df = pd.read_csv('./Geral/ArquivosGeralx.csv', sep=';', encoding='utf8')

            # concatena os DataFrame e arquiva os htm
    frame = pd.concat([frame, df])


    # Salva o arquivo atualizado
frame.to_csv('./Geral/CSV_Geral.csv', sep=',', index=False)

