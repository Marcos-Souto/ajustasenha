# -*- coding: utf-8 -*-
"""
Created on Mon Oct  1 10:42:18 2018

@author: marcos.souto
"""
import pandas as pd
import os
import glob
from log import logger
import shutil
import datetime
import Email_Senhas

# --- Frunção que le arquivos originais no diretorio -------------------------#
def LerArqGeral(strDiretorio,strArquivo,strExtecao):
    # -- Seta o diretorio
    os.chdir(strDiretorio)
    # -- Verifica se o diretorio esta vazio
    if glob.glob == []:
        logger.info("Não existe arquivos da Energisa para serem gravado")
    else:
        # -- Abre o arquivo com todo o historico para ser completado
        dados = pd.read_csv(strArquivo, sep=';', encoding='ISO-8859-1')

        # -- Realiza loop no diretorio procurando arquivo htm --#
        for files in glob.glob(strExtecao):
            try:
                # -- le o arquivo .htm e carrega para um DataFrame
                tabela = pd.read_html(files, encoding='ISO-8859-1')
                Tabela = pd.DataFrame(tabela[0])

                dados = pd.concat([dados, Tabela])
                shutil.move(files, './Original')
                logger.info("Arquivo " + files + " gravado com sucesso")
                print("Arquivo " + files + " gravado com sucesso")
            except Exception:
                # -- Grava o log para acompanhamento
                logger.error("Erro ao gravar o arquivo " + files)
                print("Erro ao gravar o arquivo " + files)

        # Salva o arquivo atualizado
        dados.to_csv('./Corrigido/Geral.csv', sep=';', encoding='ISO-8859-1')
# ----------------------------------------------------------------------------#

# -- Função interna para cria data de verificação ----------------------------#
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
            with open('../refData.txt', 'w') as arq:
                arq.write(str(dtREF))
        else:
            dtREF = datetime.datetime(refData)
    except Exception:
        Ref = now - datetime.timedelta(days=2)
        dtREF = Ref.date()
        with open('../refData.txt', 'w') as arq:
            arq.write(str(dtREF))
    return dtREF
# ----------------------------------------------------------------------------#

# --Função Que havalia se existe senha para os beneficiarios ou procedimentos-# 
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
    df = pd.read_csv(tbSenhas, sep=';', encoding='ISO-8859-1')
    dfCluster = pd.read_excel(tbCluster, sheet_name=SheetCluster)
    dfProced = pd.read_excel(r'\\GcRJ01T-SRV03\Privado\Gestão de Risco e '
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
                              row2.NOME_PROCEDIMENTO,
                              row2.TIPO_ASSOCIADO])
    # --- Verifica se encontrou alguma senha e envia e-mail
    if not lista:
        # -- Seta a menssagem do corpo do email caso não encontre senha
        msg = str("<p>E-mail enviado pela gest&atilde;o de senhas Grupo Case."
                  "</p><p> N&atilde;o foram encontradas senhas para os "
                  "benefici&aacute;rios flegados como alerta.</p>")
        # -- Chama a funçõa que envia o email sem anexo
        Email_Senhas.email_sem_anexo(msgX=msg,
                                     strTitulo=strTitulo)
        logger.info(strTitulo + " email eviado sem anexo")
        print(strTitulo + " email eviado sem anexo")

    else:
        # -- Caso encontre alguma senha monta o arquivo
        dfSenhas = pd.DataFrame(lista)# -- Carrega o data frame
        # -- Renomeio a coluna do data frame
        dfSenhas.columns = ['NUM_ASSOCIADO',
                            'NOME_ASSOCIADO',
                            'DT_OCORRENCIA',
                            'NOME_PRESTADOR',
                            'NOME_TRATAMENTO',
                            'NOME_PROCEDIMENTO',
                            'classificacao']
        # -- Grava o data frame em um aquivo xls para ser enviado como anexo
        dfSenhas.to_excel('../Anexo_Senha.xls')

        # -- Seta a menssagem do copo do email
        msg = str("<p>E-mail enviado pela gest&atilde;o de senhas Grupo Case."
                  "</p><p> No anexo est&atilde;o as senhas do ultimo dois dias"
                  "dos benefici&aacute;rios flegados para alerta.</p>")
        # -- Chama a função que envia o email com anexo
        Email_Senhas.adiciona_anexo(msgX=msg,
                                    filename='../Anexo_Senha.xls',
                                    strTitulo=strTitulo)
        logger.info(strTitulo + " email eviado com anexo de procedimento")
        print(strTitulo + " email eviado com anexo de procedimento")
