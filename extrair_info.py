import xml.dom.minidom as mnd

def extrair_info(path, roo_t, child1, child2, *argv):
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
               
    return dic

#Os seguintes prints são apenas para exibição dos dicionários.

eventos = extrair_info("curriculo.xml", 'TRABALHO-EM-EVENTOS', 'DADOS-BASICOS-DO-TRABALHO', 'DETALHAMENTO-DO-TRABALHO', 'TITULO-DO-TRABALHO', 'ANO-DO-TRABALHO', 'TITULO-DOS-ANAIS-OU-PROCEEDINGS', 'DOI')
print('-'*90,'\nTRABALHOS PUBLICADOS EM EVENTOS:\n', eventos)

artigos = extrair_info('curriculo.xml','ARTIGO-PUBLICADO','DADOS-BASICOS-DO-ARTIGO', 'DETALHAMENTO-DO-ARTIGO', 'TITULO-DO-ARTIGO', 'ANO-DO-ARTIGO', 'TITULO-DO-PERIODICO-OU-REVISTA', 'DOI')
print('-'*90,'\nTRABALHOS PUBLICADOS EM PERIÓDICOS CIENTÍFICOS:\n', artigos)

capitulos = extrair_info('curriculo.xml', 'CAPITULO-DE-LIVRO-PUBLICADO', 'DADOS-BASICOS-DO-CAPITULO', 'DETALHAMENTO-DO-CAPITULO', 'TITULO-DO-CAPITULO-DO-LIVRO', 'ANO', 'TITULO-DO-LIVRO', 'DOI')
print('-'*90,'\nCAPÍTULOS DE LIVROS PUBLICADOS:\n', capitulos)

apresentacoes = extrair_info('curriculo.xml', 'APRESENTACAO-DE-TRABALHO', 'DADOS-BASICOS-DA-APRESENTACAO-DE-TRABALHO', 'DETALHAMENTO-DA-APRESENTACAO-DE-TRABALHO', 'TITULO', 'ANO', 'NOME-DO-EVENTO', 'CIDADE-DA-APRESENTACAO', 'DOI')
print('-'*90,'\nAPRESENTAÇÕES DE TRABALHOS:\n', apresentacoes)