import xml.dom.minidom as mnd
import pandas as pd
from pandas import ExcelWriter
import xml.etree.ElementTree as ET
import os

def gerar_artigos(path, roo_t, child1, child2, *argv):
    doc = mnd.parse(path);
    contador = -1

    root = doc.getElementsByTagName(roo_t); #Pega o root dos currículos
    
    child_dadosBasicos = []; #Cria uma lista para o endereço de todas as child que serão inseridas aqui e na lista posterior
    child_detalhamento = [];

    atributos = [] #Cria uma lista para armazenar o atributo de todas child
    atributos_Tag = [] #Lista para armazenar a Tag dos atributos
    
    dic = {}#Dicionário que será retornado no final da função


    for i in root:

        child_dadosBasicos.append( i.getElementsByTagName(child1)) #Adiciona o endereço da respectiva child nas listas child_X

        child_detalhamento.append( i.getElementsByTagName(child2))
        
    for arg in argv:
        contador += 1 #Contador para alocar os atributos em seus devidos lugares
        atributos.append([]) #Cria uma sub-lista que irá armazenar os atributos de cada Tag separadamente
        atributos_Tag.append(arg) #Adiciona a Tag em questão na lista de Tag

        for i, j in zip (child_dadosBasicos, child_detalhamento):

            try:
                atributos[contador].append(i[0].attributes[arg].value) #Tenta buscar o atributo na primeira child, se não for encontrado parte para o segundo
            except KeyError:
                atributos[contador].append(j[0].attributes[arg].value)
    
    for i in range(len(atributos)):
        dic[atributos_Tag[i]] = (atributos[i]) #Adiciona para cada chave a tupla contendo todos elementos da sublista em questão
                                               
        
    contador = 0
    for j in range(len(atributos[0])): #Serve para a próxima função, uma vez que precisamos fazer uma iteração de todas linhas presentes no excel
        contador+=1
                
            
    return dic, contador

def excelMaker(dictionary):
    df = pd.DataFrame(dictionary[0]) #Cria o dataframe pelo pandas
    
    writer = ExcelWriter("artigos.xlsx") #Cria o arquivo excel
    
    workbook=writer.book #Cria a instância book para podermos utilizar a função
    
    formato = workbook.add_format({'text_wrap': True}) #Armazena o formato que buscamos, nesse caso de quebra de texto
    
    df.to_excel(writer, "Artigos_publicados", index=False) #É adicionado o nome do sheet e em seguida seleciona a opção de ter ou não index
    
    worksheet = writer.sheets['Artigos_publicados'] #Variável para identificar com qual sheet será trabalhado
    
    worksheet.set_column(0, dictionary[1], 20, formato) #É modificado o tamanho da coluna, selecionando de qual até qual coluna será modificado
    #No caso acima o último parâmetro passado se trata do formato de ter quebra de texto
    
    for i in range(dictionary[1]):
        
        worksheet.set_row(i,90, formato) #Diferente da definição da coluna, no set_row não existe parâmetro de início e fim para linha, apenas 
        #da linha em questão, por isso é necessário a iteração que se trata da variável contador da função anterior
        
    writer.save()
    writer.close() #Salva e finaliza a edição do arquivo
    

#Os 4 primeiros parâmetros se tratam do caminho do curriculo em questão, do nome da root, e o nome dos dois childs em que serão buscados
#os atributos, todos parâmetros após estes se trata dos atributos que desejamos inserir no excel.

curr_Analise = gerar_artigos("curriculo.xml",'ARTIGOS-PUBLICADOS','DADOS-BASICOS-DO-ARTIGO', 'DETALHAMENTO-DO-ARTIGO', 'TITULO-DO-ARTIGO', "ANO-DO-ARTIGO",
                      "TITULO-DO-PERIODICO-OU-REVISTA", "DOI")


#Recebe a variável anterior para criação do excel
excel = excelMaker(curr_Analise)






def gerar_trabalhos(path, roo_t, child1, child2, *argv):
    doc = mnd.parse(path);
    contador = -1

    root = doc.getElementsByTagName(roo_t); #Pega o root dos currículos
    
    child_dadosBasicos = []; #Cria uma lista para o endereço de todas as child que serão inseridas aqui e na lista posterior
    child_detalhamento = [];

    atributos = [] #Cria uma lista para armazenar o atributo de todas child
    atributos_Tag = [] #Lista para armazenar a Tag dos atributos
    
    dic = {}#Dicionário que será retornado no final da função


    for i in root:

        child_dadosBasicos.append( i.getElementsByTagName(child1)) #Adiciona o endereço da respectiva child nas listas child_X

        child_detalhamento.append( i.getElementsByTagName(child2))
        
    for arg in argv:
        contador += 1 #Contador para alocar os atributos em seus devidos lugares
        atributos.append([]) #Cria uma sub-lista que irá armazenar os atributos de cada Tag separadamente
        atributos_Tag.append(arg) #Adiciona a Tag em questão na lista de Tag

        for i, j in zip (child_dadosBasicos, child_detalhamento):

            try:
                atributos[contador].append(i[0].attributes[arg].value) #Tenta buscar o atributo na primeira child, se não for encontrado parte para o segundo
            except KeyError:
                atributos[contador].append(j[0].attributes[arg].value)
    
    for i in range(len(atributos)):
        dic[atributos_Tag[i]] = (atributos[i]) #Adiciona para cada chave a tupla contendo todos elementos da sublista em questão
                                               
        
    contador = 0
    for j in range(len(atributos[0])): #Serve para a próxima função, uma vez que precisamos fazer uma iteração de todas linhas presentes no excel
        contador+=1
                
            
    return dic, contador

def excelMaker(dictionary):
    df = pd.DataFrame(dictionary[0]) #Cria o dataframe pelo pandas
    
    writer = ExcelWriter("trabalhos.xlsx") #Cria o arquivo excel
    
    workbook=writer.book #Cria a instância book para podermos utilizar a função
    
    formato = workbook.add_format({'text_wrap': True}) #Armazena o formato que buscamos, nesse caso de quebra de texto
    
    df.to_excel(writer, "Trabalho_Eventos", index=False) #É adicionado o nome do sheet e em seguida seleciona a opção de ter ou não index
    
    worksheet = writer.sheets['Trabalho_Eventos'] #Variável para identificar com qual sheet será trabalhado
    
    worksheet.set_column(0, dictionary[1], 20, formato) #É modificado o tamanho da coluna, selecionando de qual até qual coluna será modificado
    #No caso acima o último parâmetro passado se trata do formato de ter quebra de texto
    
    for i in range(dictionary[1]):
        
        worksheet.set_row(i,90, formato) #Diferente da definição da coluna, no set_row não existe parâmetro de início e fim para linha, apenas 
        #da linha em questão, por isso é necessário a iteração que se trata da variável contador da função anterior
        
    writer.save()
    writer.close() #Salva e finaliza a edição do arquivo
    

#Os 4 primeiros parâmetros se tratam do caminho do curriculo em questão, do nome da root, e o nome dos dois childs em que serão buscados
#os atributos, todos parâmetros após estes se trata dos atributos que desejamos inserir no excel.
curr_Analise = gerar_trabalhos("curriculo.xml","TRABALHO-EM-EVENTOS",'DADOS-BASICOS-DO-TRABALHO', 'DETALHAMENTO-DO-TRABALHO', 'TITULO-DO-TRABALHO', 'ANO-DO-TRABALHO',
                      'TITULO-DOS-ANAIS-OU-PROCEEDINGS', 'DOI', 'CIDADE-DO-EVENTO')


#Recebe a variável anterior para criação do excel
excel = excelMaker(curr_Analise)
 



def gerar_capitulos():
    doc = mnd.parse("curriculo.xml")
    teste = doc.getElementsByTagName("CAPITULO-DE-LIVRO-PUBLICADO")
    childs1 = []
    childs2 = []
    tit = []
    ano = []
    tit_liv = []
    doi = []
    dic = {'TITULO': tit,'ANO': ano,'TITULO-DO-LIVRO': tit_liv,'DOI': doi}
    
    #print(teste)
    for i in teste:
        child1 = i.getElementsByTagName('DADOS-BASICOS-DO-CAPITULO')
        child2 = i.getElementsByTagName('DETALHAMENTO-DO-CAPITULO')
        childs1.append(child1)
        childs2.append(child2)
        
    
    for i in childs1:
        tit.append(i[0].attributes['TITULO-DO-CAPITULO-DO-LIVRO'].value)
        ano.append(i[0].attributes['ANO'].value)
        doi.append(i[0].attributes['DOI'].value)
        
    for i in childs2:
        tit_liv.append(i[0].attributes['TITULO-DO-LIVRO'].value)
    
    return dic

def excelMaker(dictionary):
    df = pd.DataFrame(dic) #Cria o dataframe pelo pandas
    
    writer = ExcelWriter("capitulos.xlsx") #Cria o arquivo excel
    
    workbook=writer.book #Cria a instância book para podermos utilizar a função
    
    formato = workbook.add_format({'text_wrap': True}) #Armazena o formato que buscamos, nesse caso de quebra de texto
    
    df.to_excel(writer, "Capitulos_Publicados", index=False) #É adicionado o nome do sheet e em seguida seleciona a opção de ter ou não index
    
    worksheet = writer.sheets['Capitulos_Publicados'] #Variável para identificar com qual sheet será trabalhado
    
    worksheet.set_column(0, 3, 20, formato) #É modificado o tamanho da coluna, selecionando de qual até qual coluna será modificado
    #No caso acima o último parâmetro passado se trata do formato de ter quebra de texto
    
    for i in range(3):
        
        worksheet.set_row(i,90, formato) #Diferente da definição da coluna, no set_row não existe parâmetro de início e fim para linha, apenas 
        #da linha em questão, por isso é necessário a iteração que se trata da variável contador da função anterior
        
    writer.save()
    writer.close() #Salva e finaliza a edição do arquivo

dic = gerar_capitulos()
#Recebe a variável anterior para criação do excel
excel = excelMaker(dic)




def gerar_apresentacoes():
    tree = ET.parse('curriculo.xml')
    root = tree.getroot()
     
    tit = []
    ano = []
    pais = []
    tev = []
    cid = []
    loc = []
    doi = []
     
    dicionario = {'TITULO': tit, 'ANO': ano,'PAIS': pais ,'NOME-DO-EVENTO': tev, 'LOCAL-DO-EVENTO': loc, 'CIDADE-DA-APRESENTACAO': cid, 'DOI': doi}
     
    for tecnico in root[2][13]:
        for apresentacao in tecnico.iter('APRESENTACAO-DE-TRABALHO'):
            for trabalho in apresentacao.iter('DETALHAMENTO-DA-APRESENTACAO-DE-TRABALHO'):
                tev.append(trabalho.attrib['NOME-DO-EVENTO'])
                loc.append(trabalho.attrib['LOCAL-DA-APRESENTACAO'])
                cid.append(trabalho.attrib['CIDADE-DA-APRESENTACAO'])
     
            for trabalho in apresentacao.iter('DADOS-BASICOS-DA-APRESENTACAO-DE-TRABALHO'):
                tit.append(trabalho.attrib['TITULO'])
                ano.append(trabalho.attrib['ANO'])
                pais.append(trabalho.attrib['PAIS'])
                doi.append(trabalho.attrib['DOI'])
     
    for i in dicionario:
        print(dicionario[i])
     
    chaves = dicionario.keys() #puxa as chaves do dicionario
     
    for key in chaves: #substitui qualquer valor vazio nas listas com a string "não consta"
        for element in range(0, len(dicionario[key])):
            if dicionario[key][element] == '':
                dicionario[key][element] = 'Não consta'
     
    with open('apresentacoes.csv', 'w', encoding= 'utf-8') as file:
        file.write('ANO, PAÍS, EVENTO, CIDADE DA APRESENTACAO, LOCAL DA APRESENTACAO, DOI, \n')
       
        for n in range(0, len(dicionario['ANO'])):
            for i in dicionario:
                file.write(dicionario[i][n] + ', ')
            file.write('\n')
            
gerar_apresentacoes()




def gerar_sheet():
    writer = ExcelWriter("sheet.xlsx") # criando arquivo
    workbook = writer.book # instanciando para uso dos métodos
    formato = workbook.add_format({'text_wrap': True})
    
    xls_trabalhos = pd.ExcelFile('trabalhos.xlsx') # carrega xlsx
    df = xls_trabalhos.parse('Trabalho_Eventos') # xlsx -> dataframe
    df.to_excel(writer, "Trabalhos em Eventos", index=False) # df -> excel
    worksheet1 = writer.sheets['Trabalhos em Eventos'] # escrevendo dados
    worksheet1.set_column(0, 4, 30, formato) # formato das colunas
    
    xls_periodicos = pd.ExcelFile('artigos.xlsx')
    df = xls_periodicos.parse('Artigos_publicados')
    df.to_excel(writer, "Trabalhos em Periódicos", index=False) 
    worksheet2 = writer.sheets['Trabalhos em Periódicos']
    worksheet2.set_column(0, 3, 30, formato)
    
    xls_capitulos = pd.ExcelFile('capitulos.xlsx')
    df = xls_capitulos.parse('Capitulos_Publicados')
    df.to_excel(writer, "Capítulos de Livros", index=False) 
    worksheet3 = writer.sheets['Capítulos de Livros']
    worksheet3.set_column(0, 3, 30, formato)
    
    df = pd.read_csv('apresentacoes.csv')
    df.to_excel(writer, "Apresentação de Trabalhos", index=True, index_label='TITULO DO TRABALHO') 
    worksheet4 = writer.sheets['Apresentação de Trabalhos']
    worksheet4.set_column(0, 6, 30, formato)
    
    writer.save() # salva e fecha o arquivo
    writer.close()
    
    os.remove("trabalhos.xlsx") # remove as planilhas originais
    os.remove("artigos.xlsx")
    os.remove("capitulos.xlsx")
    os.remove("apresentacoes.csv")
    
gerar_sheet()