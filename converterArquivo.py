import sys
import os
import comtypes.client
from pathlib import Path
from win32com import client

formatoPDF = 17

CAMINHO = Path('C:\\Python\\Converter Arquivos\\') #ALTERAR AQUI PARA O CAMINHO QUE TA OS ARQUIVOS
ARQUIVOS_EXCEL = []
ARQUIVOS_WORD = []
 
for NOME_ARQUIVO in CAMINHO.glob('*'):
    if NOME_ARQUIVO.suffix == '.docx'or NOME_ARQUIVO.suffix == '.doc':
        ARQUIVOS_WORD.append(NOME_ARQUIVO)
    elif NOME_ARQUIVO.suffix == '.xlsx' or NOME_ARQUIVO.suffix == '.xls':
        ARQUIVOS_EXCEL.append(NOME_ARQUIVO)

for GERAR_PDF in ARQUIVOS_WORD:
    try:
        print ("Convertendo Arquivo: " + str(GERAR_PDF) )

        ORIGEM_ARQUIVO = os.path.abspath(GERAR_PDF)
        print("Lendo Origem do Arquivo")

        SAIDA_ARQUIVO = os.path.splitext(GERAR_PDF)[0] + '.pdf'
        print("Lendo Saida do Arquivo")

        INSTANCIA_WORD = comtypes.client.CreateObject('Word.Application')
        print("Criando Instância do Word")

        DOCUMENTO_WORD = INSTANCIA_WORD.Documents.Open(ORIGEM_ARQUIVO)
        print("Abrindo documento Word")

        DOCUMENTO_WORD.SaveAs(SAIDA_ARQUIVO, FileFormat=formatoPDF)
        print("Salvando documento PDF")

        DOCUMENTO_WORD.Close()
        print("Fechando Instância")

        INSTANCIA_WORD.Quit()
        print("Processo concluido")

    except:
        print ("Erro ao tentar converter o arquivo: " + str(GERAR_PDF))
        

for GERAR_PDF in ARQUIVOS_EXCEL:
    try:
        print ("Convertendo Arquivo: " + str(GERAR_PDF) )
        
        ORIGEM_ARQUIVO = os.path.abspath(GERAR_PDF)
        print("Lendo Origem do Arquivo")

        SAIDA_ARQUIVO = os.path.splitext(GERAR_PDF)[0] + '.pdf'
        print("Lendo Saida do Arquivo")

        INSTANCIA_EXCEL = client.Dispatch("Excel.Application")
        print("Criando Instância do EXCEL")

        ARQUIVO_EXCEL = INSTANCIA_EXCEL.Workbooks.Open(ORIGEM_ARQUIVO)
        print("Abrindo documento EXCEL")

        GUIAS_PLANILHA = ARQUIVO_EXCEL.Worksheets[0]
        print("Lendo as Guias da Planilha")
        
        GUIAS_PLANILHA.ExportAsFixedFormat(0, SAIDA_ARQUIVO)
        print("Salvando documento PDF")

        ARQUIVO_EXCEL.Close()
        print("Fechando Instância")

        INSTANCIA_EXCEL.Quit()
        print("Processo concluido")

    except:
        print ("Erro ao tentar converter o arquivo: " + str(GERAR_PDF))
        
