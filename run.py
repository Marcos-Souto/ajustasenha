# -*- coding: utf-8 -*-
"""
Created on Mon Oct  1 15:22:16 2018

@author: marcos.souto
"""
import Gestao_de_Senhas as GS
import CapturaLefort

# -- Chama arquivos da Valid -------------------------------------------------#
# -- Seta o diretorio onde estão os arquivo csv corrigido
my_dir1 = str(r'\\GcRJ01T-SRV03\Privado\Gestão de Risco e '
              'Saúde\VALIDAÇÕES\CNU\VALID senhas\Relatorios')
GS.LerArqGeral(my_dir1, './Corrigido/Geral.csv', '*.xls')

# --chama a função que verifica, existe  benficiario ou procedimentos de alerta
GS.Ver_Senhas(
        strDiretorio=my_dir1,
        tbSenhas='./Corrigido/Geral.csv',
        tbCluster='../Resumo_Cluster_Valid.xlsx',
        SheetCluster='Classificados',
        strTitulo='Gestão de Senhas Valid'
        )

# -- Chama arquivo da Energisa -----------------------------------------------#
# -- Seta o diretorio onde estão os arquivo csv corrigido
my_dir2 = str(r'\\GcRJ01T-SRV03\Privado\Gestão de Risco e '
              'Saúde\VALIDAÇÕES\CNU\Energisa_Senhas\Relatorio')
GS.LerArqGeral(my_dir2, './Corrigido/Geral.csv', '*.xls')

# --chama a função que verifica, existe  benficiario ou procedimentos de alerta
GS.Ver_Senhas(
        strDiretorio=my_dir2,
        tbSenhas='./Corrigido/Geral.csv',
        tbCluster='../Resumo_Cluster.xlsx',
        SheetCluster='Classificados',
        strTitulo='Gestão de Senhas Energisa'
        )


# -- Chama arquivo da Paranapanema--------------------------------------------#
# -- Seta o diretorio onde estão os arquivo csv corrigido
my_dir3 = str(r'\\GcRJ01T-SRV03\Privado\Gestão de Risco e '
              'Saúde\VALIDAÇÕES\CNU\Paranapanema\Relatorios')
GS.LerArqGeral(my_dir3, './Corrigido/Geral.csv', '*.xls')

# --chama a função que verifica, existe  benficiario ou procedimentos de alerta
GS.Ver_Senhas(
        strDiretorio=my_dir3,
        tbSenhas='./Corrigido/Geral.csv',
        tbCluster='../Resumo_Cluster.xlsx',
        SheetCluster='Classificados',
        strTitulo='Gestão de Senhas Paranapanema'
        )

# -- Captura Lefort ----------------------------------------------------------#
my_dir4 = str(r'\\GcRJ01T-SRV03\Privado\Gestão de Risco e '
              'Saúde\VALIDAÇÕES\Bradesco\Lefort')
CapturaLefort.func_Descompacta(my_dir4)
CapturaLefort.func_Ler_Arquivo(my_dir4, '../Senhas_Lefort.xls')
