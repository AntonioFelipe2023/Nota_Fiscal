import xmltodict
import pandas as pd
import os
import time
import pandas as pd
import os
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located
from selenium.webdriver.support.expected_conditions import visibility_of, invisibility_of_element
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions

from selenium.webdriver.chrome.options import Options

def ler_danfe(nota):

    #Para ler um arquivo em XML
    #r'C:\Users\Felipe\PycharmProjects\Nota fiscal\Nf´s\DANFEBrota.xml','rb'

    with open(nota,'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
    #print(documento)

    #Apoio/caminho de onde todas as informações da NF vão estar
    apoio=documento['nfeProc']['NFe']['infNFe']


    #Dados necessarios
    nf=apoio['ide']['nNF']
    valor=apoio['total']['ICMSTot']['vNF']


    empresa_vendeu=apoio['emit']['xNome']
    uf_prestador=apoio['emit']['enderEmit']['UF']

    cnpj_prestador=apoio['emit']['CNPJ']



    empresa_comprou=apoio['dest']['xNome']
    uf_comprou=apoio['dest']['enderDest']['UF']

    cnpj_tomador=apoio['dest']['CPF']


    resposta = {
        'nf':nf,
        'valor':valor,
        'empresa_vendeu':empresa_vendeu,
        'uf_prestador':uf_prestador,

        'cnpj_prestador':cnpj_prestador,
        'empresa_comprou':empresa_comprou,
        'uf_comprou':uf_comprou,

        'cnpj_tomador':cnpj_tomador,

    }

    return resposta


def ler_servico(nota):

    #Para ler um arquivo em XML
    #r'C:\Users\Felipe\PycharmProjects\Nota fiscal\Nf´s\DANFEBrota.xml','rb'

    with open(nota,'rb') as arquivo:
        documento = xmltodict.parse(arquivo)
    #print(documento)

    #Apoio/caminho de onde todas as informações da NF vão estar
    apoio=documento['ConsultarNfseResposta']['ListaNfse']['CompNfse']['Nfse']['InfNfse']



    #Dados necessarios
    nf=apoio['Numero']
    valor=apoio['Servico']['Valores']['ValorServicos']


    empresa_vendeu=apoio['PrestadorServico']['RazaoSocial']
    uf_prestador=apoio['PrestadorServico']['Endereco']['Uf']

    cnpj_prestador=apoio['PrestadorServico']['IdentificacaoPrestador']['Cnpj']



    empresa_comprou=apoio['TomadorServico']['RazaoSocial']
    uf_comprou=apoio['TomadorServico']['Endereco']['Uf']

    cnpj_tomador=apoio['TomadorServico']['IdentificacaoTomador']['CpfCnpj']['Cnpj']


    resposta = {
        'nf':nf,
        'valor':valor,
        'empresa_vendeu':empresa_vendeu,
        'uf_prestador':uf_prestador,

        'cnpj_prestador':cnpj_prestador,
        'empresa_comprou':empresa_comprou,
        'uf_comprou':uf_comprou,

        'cnpj_tomador':cnpj_tomador,

    }

    return resposta




resultados = []
nomes_arquivos_pdf = []


lista_arquivos = os.listdir('Nf´s')




for arquivo in lista_arquivos:
    if 'xml' in arquivo:
        if 'DANFE' in arquivo:
            resultados.append(ler_danfe(f'Nf´s/{arquivo}'))
        else:
            resultados.append(ler_servico(f'Nf´s/{arquivo}'))




#transformar um dicionário que foi chamado de resposta em uma tabela no excel
#tabela = pd.DataFrame.from_dict([resposta])
tabela = pd.DataFrame(resultados)

tabela.to_excel(r"C:\Users\Felipe\Downloads\Notas Fiscais.xlsx",index=False)

#Criar uma coluna atraves da base que criei


arquivo_excel = r"C:\Users\Felipe\Downloads\Notas Fiscais.xlsx"
tabela_nf = pd.read_excel(arquivo_excel)

lista_arquivos = os.listdir('Nf´s')


nomes_arquivos_pdf = []


for arquivo in lista_arquivos:
    if 'pdf' in arquivo:
        nomes_arquivos_pdf.append(arquivo)

tabela_nf['nomeArquivo'] = nomes_arquivos_pdf

# Salvar a tabela atualizada de volta ao arquivo Excel
tabela_nf.to_excel(arquivo_excel, index=False)


time.sleep(2)

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)



arquivo=r"C:\Users\Felipe\PycharmProjects\Nota fiscal\Nf´s\index.html"
navegador.get(arquivo)

navegador.maximize_window()

nf='nf'
valor='valor'
prestador='empresa_vendeu'
ufprestador='uf_prestador'
cnpjprestador='cnpj_prestador'
tomador='empresa_comprou'
uftomador='uf_comprou'
cnpjtomador='cnpj_tomador'
anexo='file'
enviar='/html/body/form/input[10]'


wait = WebDriverWait(navegador,120)
wait.until(expected_conditions.visibility_of_element_located((By.ID,nf)))


elemento=navegador.find_element(By.ID, 'nf')
elemento.click()



tabela= pd.read_excel(r"C:\Users\Felipe\Downloads\Notas Fiscais.xlsx")

for linha in tabela.index:
    nf=tabela.loc[linha,'nf']
    valor = tabela.loc[linha, 'valor']
    empresa_vendeu = tabela.loc[linha, 'empresa_vendeu']
    uf_prestador = tabela.loc[linha, 'uf_prestador']
    cnpj_prestador = tabela.loc[linha, 'cnpj_prestador']
    empresa_comprou = tabela.loc[linha, 'empresa_comprou']
    uf_comprou = tabela.loc[linha, 'uf_comprou']
    cnpj_tomador = tabela.loc[linha, 'cnpj_tomador']
    nomeArquivo = tabela.loc[linha, 'nomeArquivo']
    navegador.find_element(By.ID, 'nf').send_keys(str(nf))
    navegador.find_element(By.ID, 'valor').send_keys(str(valor))
    navegador.find_element(By.ID, 'empresa_vendeu').send_keys(str(empresa_vendeu))
    navegador.find_element(By.ID, 'uf_prestador').send_keys(str(uf_prestador))
    navegador.find_element(By.ID, 'cnpj_prestador').send_keys(str(cnpj_prestador))
    navegador.find_element(By.ID, 'empresa_comprou').send_keys(str(empresa_comprou))

    navegador.find_element(By.ID, 'uf_comprou').send_keys(str(uf_comprou))
    navegador.find_element(By.ID, 'cnpj_tomador').send_keys(str(cnpj_tomador))

    caminho1=r"C:\Users\Felipe\PycharmProjects\Nota fiscal\Nf´s\\"
    navegador.find_element(By.ID, 'file').send_keys(caminho1+nomeArquivo)
    time.sleep(3)
    navegador.find_element(By.XPATH, '/html/body/form/input[10]').click()
    time.sleep(2)
    




navegador.quit()