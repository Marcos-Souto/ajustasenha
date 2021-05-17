

import pandas as pd
import os
import glob
from datetime import date
import shutil
from log import logger


# -- Seta o diretorio
my_dir = str(r'\\GcRJ01T-SRV03\Privado\Gestão de Risco e Saúde\VALIDAÇÕES\UnimedRio\Arquivos')
os.chdir(my_dir)

# -- Abre o arquivo com todo o historico para ser completado
frame = pd.read_csv('ArquivosRIP.csv', sep=';', encoding='utf8')

# -- Verifica se o diretorio esta vazio
tx = os.path.exists("*.htm")
if tx == False:
    logger.info("Não existe arquivos da InimedRio para serem gravados")

# -- Realiza loop no diretorio procurando arquivo htm --#
for files in glob.glob('*.htm'):
    try:
        # -- le o arquivo .htm e carrega para um DataFrame
        tabela = pd.read_html(files)
        frm = pd.DataFrame(tabela[1])

        df = frm.transpose()  # -- Cria o DataFrame do tranposto
        df.columns = [df.loc[0]]  # Usa a primeira linha para nomear as colunas
        df.drop([0, 1], inplace=True)  # Retira a primeira e a segunda linha
        df = df.dropna(axis=1, how='all')  # Retira as colunas que não tem valor

        # -- Salva o DataFrame como CSV e depois abre ---
        df.to_csv('ArquivosRIPx.csv', sep=',', index=False)
        df1 = pd.read_csv('./ArquivosRIPx.csv', sep=',', encoding='utf8')

        # Inclui uma colona com o data atual
        df1["Data"] = date.today()
        # concatena os DataFrame e arquiva os htm
        frame = pd.concat([frame, df1])
        shutil.move(files, './ArqFinali')
        # Grava o log
        logger.info("Arquivo " + files + " gravado com sucesso")
    except:
        logger.error("Erro ao gravar o arquivo " + files)

# Salva o arquivo atualizado
frame.to_csv('./ArquivosRIP.csv', sep=';', index=False)
