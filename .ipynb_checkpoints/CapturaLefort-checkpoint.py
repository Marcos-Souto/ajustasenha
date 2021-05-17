# -*- coding: utf-8 -*-
"""
Created on Mon Nov  5 10:43:54 2018

@author: marcos.souto
"""
import zipfile
import os
import shutil
import glob
from log import logger
import pandas as pd
from datetime import datetime
import re

# - Função quele arquivo Zip e extrai conteudo
def func_Descompacta(strDiretorio):
    # -- Seta o diretorio
    os.chdir(strDiretorio)
    # -- Verifica se o diretorio esta vazio
    if glob.glob == []:
        logger.info("Não existe arquivos da Energisa para serem gravado")

    else:
        # -- Realiza loop no diretorio procurando arquivo htm --#
        for files in glob.glob('*.zip'):
            try:
                # - ler arquivo zipado e extrai conteudo
                Arq_zip = zipfile.ZipFile(files)
                Arq_zip.extractall(pwd='72459'.encode('cp850', 'replace'))
                Arq_zip.close()
                shutil.move(files, './zip_Original')

                # -Extrai o conteudo do arquivo zipado e renomeia
                nome = str('Extrct' + str(datetime.now()))
                nome = re.sub(':', '-', nome[:-7:]) + '.xls'
                os.rename('Relatórios de Estipulantes -  Bandeirantes .xlsx',
                          nome)

                # - Grava o log da tarefa
                logger.info("Arquivo " + files + " gravado com sucesso")
                print("Arquivo " + files + " gravado com sucesso")

            except Exception:
                # -- Grava o log para acompanhamento
                logger.error("Erro ao descompactar o arquivo " + files)
                print("Erro ao descompactar o arquivo ")


def func_Ler_Arquivo(strDir, strArquivo):
    os.chdir(strDir)
    # -- Verifica se o diretorio esta vazio
    if glob.glob == []:
        logger.info("Não existe arquivos da Energisa para serem gravado")
    else:
        # -- Abre o arquivo com todo o historico para ser completado
        Arq = pd.read_excel(strArquivo)
        # -- Realiza loop no diretorio procurando arquivo htm --#
        for files in glob.glob('*.xls'):
            try:
                tabela = pd.read_excel(files, encoding='ISO-8859-1')
                
                # - Renomeia colunas do DataFrame
                df = tabela.set_axis(['Estipulante',
                                      'x1',
                                      'Solicitação',
                                      'Senha',
                                      'dtSolicita',
                                      'x2',
                                      'Status',
                                      'Prorroga',
                                      'x3',
                                      'x4',
                                      'Tipo_Procedimento',
                                      'x5',
                                      'x6',
                                      'Cod_Proced',
                                      'Desc_Proced',
                                      'x7',
                                      'Apolice',
                                      'Cod_Prestador',
                                      'Desc_Prestados',
                                      'Cod_Beneficiario',
                                      'Nome',
                                      'Tipo_Internação',
                                      'Quant_Senha'],
                    axis='columns', inplace=False)

                # - Filtra linhas sem valores e dropa colunas desnecessarias
                df = df.query('Estipulante == "HOSPITAL BANDEIRANTES SA"')
                df = df.drop(columns=['x1',
                                      'x2',
                                      'x3',
                                      'x4',
                                      'x5',
                                      'x6',
                                      'x7'])

                # -Move o arquivo lido para a pasta
                shutil.move(files, './Arquivos')

                # - Grava o log da tarefa
                logger.info("Arquivo " + files + " gravado com sucesso")
                print("Arquivo " + files + " gravado com sucesso")

                # - Concatena os arquivos novos com os anctigos
                Arq = pd.concat([Arq, df])
            except Exception:
                # -- Grava o log para acompanhamento
                logger.error("Erro ao descompactar o arquivo " + files)
                print("Erro ao descompactar o arquivo ")

    Arq.to_excel(strArquivo)
