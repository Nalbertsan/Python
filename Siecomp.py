import pandas as pd

class Pessoa:
    nome = ""
    matricula = ""
    Lista = []

    def __init__(self, nome, matricula):
        self.nome = nome
        self.matricula = matricula

    def addLista(self, presenca):
        self.Lista = presenca 
    
    
    def __eq__(self, other):
        return  self.matricula == other.matricula
    

def cadastro(nome, matricula):
    return Pessoa(nome, matricula)

def percorreLista(tabela,listapessoas,listaApresentacao):
    for i in range(len(tabela)):
        listaVazia = []
        pessoa = cadastro(tabela["Nome Completo"][i], tabela["Número de matrícula"][i])
        if(tabela["Qual palestra/oficina está participando?"][i] not in listaApresentacao):
            listaApresentacao.append(tabela["Qual palestra/oficina está participando?"][i])
        if pessoa not in listapessoas:
            listaVazia.append(tabela["Qual palestra/oficina está participando?"][i])
            pessoa.addLista(listaVazia)
            listapessoas.append(pessoa)
        else:
            listapessoas[listapessoas.index(pessoa)].Lista.append(tabela["Qual palestra/oficina está participando?"][i])

def imprimeLista(listapessoas):
    print(f"Lista de pessoas: Total = {len(listapessoas)}")
    for i in range(len(listapessoas)):
        print("Nome: ",listapessoas[i].nome)
        print("Matricula: ",listapessoas[i].matricula)
        print("Lista: ", listapessoas[i].Lista)

def transformar(lista):
    dic = {}
    listNome = []
    listMatricula = []
    listPresenca = []
    for pessoa in lista:
        listNome.append(pessoa.nome)
        listMatricula.append(pessoa.matricula)
        listPresenca.append(pessoa.Lista)
    dic["Nome"] = listNome
    dic["Matricula"] = listMatricula
    dic["Soma"] = somar(listPresenca)
    return dic

def somar(lista):
    listasoma = []
    for i in range(len(lista)):
        soma = 0
        for list in set(lista[i]):
            if(list == "INTRODUÇÃO A METODOLOGIA PBL"):
                soma+=9
            elif(list == "SELEC * FROM ECOMP"):
                soma+=2
            elif(list == "CONHECENDO O RAMO ESTUDANTIL"):
                soma+=2
            elif(list == "O INTERCÂMBIO NA UEFS - O PAPEL DA AERI"):
                soma+=1
            elif(list == "INTRODUÇÃO AO LABORATÓRIO PARA CIRCUITOS ELÉTRICOS"):
                soma+=5
            elif(list == "TRAJETÓRIA DO ENGENHEIRO DE SOFTWARE"):
                soma+=1
            elif(list == "PALESTRA SOBRE ENADE"):
                soma+=1
            elif(list == "CURRICULARIZAÇÃO DA EXTENSÃO NO CURSO"):
                soma+=2
            elif(list == "MATEMÁTICA & PROSA"):
                soma+=2
            elif(list == "MATEMÁTICA DISCRETAS E SUAS APLICAÇÕES"):
                soma+=2
            elif(list == "ECOMPLICADO FAZER ECOMP?"):
                soma+=2
            elif(list == "EXPLICANDO TODAS AS MATÉRIAS DO CURSO"):
                soma+=2
        if(soma > 20):
            soma = 20
        if(soma == 4):
            soma = 5
        listasoma.append(soma)
    return listasoma

def transformar_Dic_cadaPre(lista, listaApresentacao):
    dic = {}
    for apresentacao in listaApresentacao:
        dic[apresentacao] = []
        for pessoa in lista:
            if apresentacao in pessoa.Lista:
                dic[apresentacao].insert(0, pessoa.nome)
            else:
                dic[apresentacao].append("")
    return dic

def main():
    listaApresentacao = []
    listapessoas = []
    tabela = pd.read_excel("Python\Lista.xlsx")
    percorreLista(tabela, listapessoas,listaApresentacao)
    ##imprimeLista(listapessoas)
    dic = transformar(listapessoas)
    dic2 = transformar_Dic_cadaPre(listapessoas, listaApresentacao)
    df = pd.DataFrame(dic)
    df.to_excel("Python\ListaNumeros.xlsx", index=False)
    #df2 = pd.DataFrame(dic2)
    #df2.to_excel("Python\ListaOrganizadaPalestras.xlsx", index=False)
    #df.to_excel("Python\ListaOrganizada.xlsx", index=False)
    

main()


