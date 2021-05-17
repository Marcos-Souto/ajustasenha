# -*- coding: utf-8 -*-
"""
Created on Wed Sep 26 14:07:05 2018

@author: marcos.souto
"""

import pandas as pd
import os
import glob


def LerCVSGeral(strDiretorio):
    # -- Seta o diretorio
    os.chdir(strDiretorio)

    # -- Abre o arquivo com todo o historico para ser completado
    frame = pd.read_csv('./Geral/CSV_Geralx.csv', sep=',', encoding='utf8')

    # -- Realiza loop no diretorio procurando arquivo htm --#
    for files in glob.glob('*.csv'):
        # -- le o arquivo .htm e carrega para um DataFrame
        tabela = pd.read_csv(files, sep=';',
                             error_bad_lines=False,
                             encoding='utf8')
        frame = pd.concat([frame, tabela])

    # Salva o arquivo atualizado
    frame.to_csv('./Geral/CSV_Geral.csv', sep=';', encoding='utf8')
