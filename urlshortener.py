from mailmerge import MailMerge
import pandas as pd
from docx2pdf import convert
import pathlib
import requests
import json

# Configuração


f = open("config.txt", "r")
accessToken = f.readline().strip()
group_guid = f.readline().strip()
f.close()
postUrl = "https://api-ssl.bitly.com/v4/shorten"

# Funções


def shortenUrl(urlToBeShortened):
    objeto = {
        'group_guid': group_guid,
        'domain': 'bit.ly',
        'long_url': urlToBeShortened
    }
    headers = {'Content-Type': 'application/json',
               'Authorization': accessToken}
    x = requests.post(postUrl, json=objeto, headers=headers)
    return x.json()['link']


def inicializarMateriasAlunoSemCorrespondencia():
    todasAsMaterias = []
    for _ in range(0, 4):
        materia = "Estudo livre."
        todasAsMaterias.append(materia)
    return todasAsMaterias


def obterPastaAtual():
    return pathlib.Path(__file__).parent.absolute()


def converterDocumentosParaPDF():
    pastaAtual = obterPastaAtual()
    convert(pastaAtual, pastaAtual)


def checaSeMateriaCorresponde(nomeDoAluno, nomeDoAlunoPlanilhaDificuldades):
    houveCorrespondencia = False

    houveCorrespondencia = True
    return houveCorrespondencia


def gerarDocx(nomeDoAluno, todasAsMaterias):
    documentoDeSaida = MailMerge(templateDeSaida)
    documentoDeSaida.merge(nomeAluno=nomeDoAluno, sabado01=todasAsMaterias[0],
                           sabado02=todasAsMaterias[1], domingo01=todasAsMaterias[2], domingo02=todasAsMaterias[3])
    documentoDeSaida.write('GRADE - ' + nomeDoAluno + '.docx')


def gerarDocumentosDeSaida():
    for _, linha in planilhaHorario.iterrows():
        nomeDoAluno = str(linha[0])
        todasAsMaterias = []
        houveCorrespondencia = False
        for _, linhaPlanilhaDificuldades in planilhaDificuldades.iterrows():
            nomeDoAlunoPlanilhaDificuldades = str(linhaPlanilhaDificuldades[0])
            if (nomeDoAluno.lower().strip() == nomeDoAlunoPlanilhaDificuldades.lower().strip()):
                houveCorrespondencia = True
                todasAsMaterias = obterTodasAsMateriasComDificuldade(
                    linhaPlanilhaDificuldades, nomeDoAluno)
        if not houveCorrespondencia:
            alunosSemCorrespondencia.append(nomeDoAluno)
            todasAsMaterias = inicializarMateriasAlunoSemCorrespondencia()
        gerarDocx(nomeDoAluno, todasAsMaterias)


print("Início da execução.")

urlReduzida = shortenUrl('https://www.fast.com')
print (urlReduzida)

print("Fim da execução.")
