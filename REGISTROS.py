# CÓDIGO PARA EXTRAIR DADOS DE PDF DE REGISTRO#
import os
import pdfplumber
import pandas as pd

path='C:\\Users\\Julio\\Desktop\\Abner\\Ficha de Registro'
os.chdir(path)
list=os.listdir(path)
print(list)
colab=[]
for j in range(len(list)):
    pdf= pdfplumber.open(list[j])
    print(list[j])
    n=range(len(pdf.pages))
    
    for i in n:
        page=pdf.pages[i]
        text=page.extract_text()
        Empresa= text.split('\n')[2].split(':')[1]
        CNPJ= text.split('\n')[2].split(':')[2]
        CEP_empresa= text.split('\n')[3].split(':')[2]
        End_empresa= text.split('\n')[4].split(':')[1].split('CÓDIGO')[0]
        Cidade= text.split('\n')[5].split(':')[1].split('CNAE')[0]
        Cnae= text.split('\n')[5].split(':')[2]
        Nome= text.split('\n')[8].split(':')[1].split('NR')[0]
        N_registro= text.split('\n')[8].split(':')[2]
        Mae_nome= text.split('\n')[9].split(':')[1].split('ESTADO')[0]
        Pai_nome= text.split('\n')[10].split(':')[1].split('GRAU')[0]
        Estado_civil= text.split('\n')[9].split(':')[2]
        Grau_instrucao= text.split('\n')[10].split(':')[2]
        Endereco= text.split('\n')[11].split(':')[1].split('SEXO')[0]
        bairro= text.split('\n')[13].split(':')[1].split('NACIONALIDADE')[0]
        cidade= text.split('\n')[12].split(':')[1].split('DATA')[0]
        cep= text.split('\n')[14].split(':')[1].split('ANO')[0]
        sexo= text.split('\n')[11].split(':')[2]
        Nacionalidade= text.split('\n')[13].split(':')[2]
        celular= text.split('\n')[15].split(':')[1].split('CIDADE')[0]
        cidade_nascimento= text.split('\n')[15].split(':')[2]
        UF_nascimento= text.split('\n')[16].split(':')[2]
        celular=text.split('\n')[16].split(':')[1].split('UF')[0]
        CPF= text.split('\n')[18].split(':')[1].split('CTPS')[0]
        RG= text.split('\n')[19].split(':')[1].split('TÍTULO')[0]
        data_nasc= text.split('\n')[12].split(':')[2]
        CTPS= text.split('\n')[18].split(':')[2]
        TITULO_eleitor= text.split('\n')[19].split(':')[2]
        PIS_PASEP_EMISSAO= text.split('\n')[20].split(':')[1].split('CERTIFICADO')[0]
        cert_militar= text.split('\n')[20].split(':')[2]
        CNH= text.split('\n')[21].split(':')[1]
        data_admissao= text.split('\n')[25].split(':')[1].split('FUNÇÃO')[0]
        funcao= text.split('\n')[25].split(':')[2]
        salario_admss= text.split('\n')[26].split(':')[1].split('CBO')[0]
        horas_semanais= text.split('\n')[28].split(':')[1]

        colaborador=[Nome, celular,  CPF, RG, CTPS, CNH, data_admissao, salario_admss, horas_semanais, TITULO_eleitor, data_nasc,PIS_PASEP_EMISSAO, funcao, N_registro, Mae_nome, Pai_nome, Estado_civil, Grau_instrucao, cert_militar,  Endereco, bairro, cidade, cep, sexo, cidade_nascimento, UF_nascimento,  Nacionalidade, celular, Empresa, CNPJ, CEP_empresa, End_empresa, Cidade, Cnae]          

#CRIA LISTA COM DADOS DE TODOS OS COLABORADORES#
    
        colab.append(colaborador)
    
#CRIA DATAFRAME COM DADOS#

dataframe= pd.DataFrame(colab, columns= ['Nome', 'celular',  'CPF', 'RG', 'CTPS', 'CNH', 'data_admissao', 'salario_admss', 'horas_semanais', 'TITULO_eleitor', 'data_nasc', 'PIS_PASEP_EMISSAO', 'funcao', 'N_registro', 'Mae_nome', 'Pai_nome', 'Estado_civil', 'Grau_instrucao', 'cert_militar',  'Endereco', 'bairro', 'cidade', 'cep', 'sexo', 'cidade_nascimento', 'UF_nascimento',  'Nacionalidade', 'celular', 'Empresa', 'CNPJ', 'CEP_empresa', 'End_empresa', 'Cidade', 'Cnae'])

#SALVA DATAFRAME COMO PLANILHA NA PASTA ABNER#
os.chdir('C:\\Users\\Julio\\Desktop\\Abner\\planilhas')
dataframe.to_excel('funcionários.xlsx')
