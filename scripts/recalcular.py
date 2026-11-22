# openpyxl grava as fórmulas sem resultado. Quem calcula é o Excel instalado.
import re
import sys
import zipfile
from pathlib import Path

import win32com.client


def calcular(caminho):
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    try:
        livro = excel.Workbooks.Open(str(caminho))
        excel.CalculateFull()
        livro.Worksheets("Painel").Activate()
        livro.Save()
        livro.Close(False)
    finally:
        excel.Quit()


def tirar_dados_pessoais(caminho):
    # ao salvar, o Excel grava a pasta do computador e o usuário do Windows dentro do arquivo
    with zipfile.ZipFile(caminho) as origem:
        partes = [(item.filename, origem.read(item.filename)) for item in origem.infolist()]
    with zipfile.ZipFile(caminho, "w", zipfile.ZIP_DEFLATED) as destino:
        for nome, dados in partes:
            if nome == "xl/workbook.xml":
                bloco = rb"<mc:AlternateContent[^>]*><mc:Choice[^>]*><x15ac:absPath.*?</mc:AlternateContent>"
                dados = re.sub(bloco, b"", dados, flags=re.DOTALL)
            elif nome == "docProps/core.xml":
                dados = re.sub(rb"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", b"", dados)
            destino.writestr(nome, dados)


if __name__ == "__main__":
    arquivo = Path(sys.argv[1]).resolve()
    calcular(arquivo)
    tirar_dados_pessoais(arquivo)
