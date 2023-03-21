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
    dic["Presenca"] = listPresenca
    return dic

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
    #df = pd.DataFrame(dic)
    df2 = pd.DataFrame(dic2)
    df2.to_excel("Python\ListaOrganizadaPalestras.xlsx", index=False)
    #df.to_excel("Python\ListaOrganizada.xlsx", index=False)
    

main()


