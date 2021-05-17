# -*- coding: utf-8 -*-
"""
Created on Tue Sep 25 10:47:28 2018

@author: marcos.souto
"""

import pandas as pd
import os
import datetime
import Email_Senhas


def Cria_Data(diretorio):
    # --- cria data de refecencia para filtrar base de senhas ---------------#
    now = datetime.datetime.now()
    my_dir = str(diretorio)
    os.chdir(my_dir)
    try:
        with open('./refData.txt', 'r') as arq:
            refData = arq.readlines()
        if now.hour < 12:
            dtREF = datetime.datetime(refData) + datetime.timedelta(days=1)
            with open('./refData.txt', 'w') as arq:
                arq.write(str(dtREF))
        else:
            dtREF = datetime.datetime(refData)
    except Exception:
        Ref = now - datetime.timedelta(days=2)
        dtREF = Ref.date()
        with open('./refData.txt', 'w') as arq:
            arq.write(str(dtREF))
    return dtREF
    # -------------------------------------------------------------------------


def Ver_Senhas(strDiretorio, tbSenhas, tbCluster, SheetCluster, strTitulo):
    """
    strDiretorio = diretorio onde se encontra as planilhas
    tbSenha = nome da planilha excel com todas as senhas recebidas => Exemplo ./Senhas_Valid.xlsm
    SheetSenha = nome da aba da planilha onde estão as senhas => Exemplo Geral
    tbCluster = nome da planilha excel com todos os beneficiarios cluterizados => Exemplo ./Resumo_Cluster_Valid.xlsx
    ShhetCluster = nome da aba da planilha onde estão  os beneficiarios => Exemplo Classificação
    """
    dtREF = Cria_Data(strDiretorio)
    my_dir = str(strDiretorio)
    os.chdir(my_dir)


    # ---  Abre as bases e carrega para do DataFrame -------------------------
    df = pd.read_csv(tbSenhas, sep=';', encoding='utf8')
    dfCluster = pd.read_excel(tbCluster, sheet_name=SheetCluster)
    dfProced = pd.read_excel(r'\\GcRJ01T-SRV03\Privado\Gestão de Risco e '\
                             'Saúde\VALIDAÇÕES\CNU\TabelaCBHPMv5.xlsx',
                             sheet_name='TabelaCBHPM')

    df['DT_OCORRENCIA'] = pd.to_datetime(df['DT_OCORRENCIA'], format='%d/%m/%Y')

    # --- Filtra os DataFrame ------------------------------------------------
    dfFitrado = df.loc[lambda df: df.DT_OCORRENCIA > dtREF, :]
    dfClusteFiltrado = dfCluster.loc[lambda dfCluster:dfCluster.classificacao != 1, :]
    dfPocedFiltrado = dfProced.loc[lambda dfProced:dfProced.Avaliar == 'x', :]
    # --- loop para verificar se existe senhas dos beneficiario flegados
    lista = []
    for row1 in dfClusteFiltrado.itertuples():
        for row2 in dfFitrado.itertuples():
            if row2.NOME_ASSOCIADO == row1.Nome_Beneficiario:
                lista.append([row2.NUM_ASSOCIADO,
                              row2.NOME_ASSOCIADO,
                              row2.DT_OCORRENCIA,
                              row2.NOME_PRESTADOR,
                              row2.NOME_TRATAMENTO,
                              row2.NOME_PROCEDIMENTO,
                              row1.classificacao])
    # --- loop para verificar se existe senhas dos procedimentos flegados
    for row1 in dfPocedFiltrado.itertuples():
        for row2 in dfFitrado.itertuples():
            if row1.Cód_Procedimento == row2.COD_PROCEDIMENTO:
                lista.append([row2.NUM_ASSOCIADO,
                              row2.NOME_ASSOCIADO,
                              row2.DT_OCORRENCIA,
                              row2.NOME_PRESTADOR,
                              row2.NOME_TRATAMENTO,
                              row2.NOME_PROCEDIMENTO])
    # --- Verifica se encontrou alguma senha e envia e-mail
    if not lista:
        # -- Seta a menssagem do corpo do email caso não encontre senha 
        msg = "<p>E-mail enviado pela gest&atilde;o de senhas Grupo Case."\
            "</p><p> N&atilde;o foram encontradas senhas para os "\
            "benefici&aacute;rios flegados como alerta.</p>"
        # -- Chama a funçõa que envia o email sem anexo
        Email_Senhas.email_sem_anexo(msgX=msg,
                                     strTitulo=strTitulo)

    else:
        # -- Caso encontre alguma senha monta o arquivo
        dfSenhas = pd.DataFrame(lista) # -- Carrega o data frame
        # -- Renomeio a coluna do data frame
        dfSenhas.columns = ['NUM_ASSOCIADO',
                            'NOME_ASSOCIADO',
                            'DT_OCORRENCIA',
                            'NOME_PRESTADOR',
                            'NOME_TRATAMENTO',
                            'NOME_PROCEDIMENTO',
                            'classificacao']
        # -- Grava o data frame em um aquivo xls para ser enviado como anexo
        dfSenhas.to_excel('./Anexo_Senha.xls')

        # -- Seta a menssagem do copo do email
        msg = "<p>E-mail enviado pela gest&atilde;o de senhas Grupo Case.</p>"\
            "<p> No anexo est&atilde;o as senhas do ultimo dois dias dos "\
            "benefici&aacute;rios flegados para alerta.</p>"
        # -- Chama a função que envia o email com anexo
        Email_Senhas.adiciona_anexo(msgX=msg,
                                    filename='./Anexo_Senha.xls',
                                    strTitulo=strTitulo)
