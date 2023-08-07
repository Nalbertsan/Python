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
    return Pessoa(nome.title(), matricula)

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
            if(list == "Quest"):
                soma+=4
            elif(list == "IEEE UEFS"):
                soma+=2
            elif(list == "Sistemas de Controle"):
                soma+=1
            elif(list == "NAPP"):
                soma+=2
            elif(list == "Apresentação geral do curso de Engenharia de Computação"):
                soma+=1
            elif(list == "Introdução à metodologia PBL"):
                soma+=8
            elif(list == "Tenho uma ideia inovadora, O que fazer?"):
                soma+=2
            elif(list == "A trajetória do engenheiro de software"):
                soma+1
            elif(list == "Mesa redonda Sobre IA e o futuro"):
                soma+=2
            elif(list == "Operação Zamazenta"):
                soma+=1
            elif(list == "Estágio não obrigatório versus Estágio Obrigatório"):
                soma+=2
            elif(list == "Ras Apresenta: Introdução a circuitos digitais"):
                soma+=4

            elif(list == "Explicando todas as matérias do curso"):
                soma+=2
            elif(list == "Alem da sala de aula: A importância das atividades complementares"):
                soma+=2
            elif(list == "Como funciona a UEFS e o que voce tem a ver com isso?"):
                soma+=1
            elif(list == "AERI"):
                soma+=1
            
            elif(list == "Carreira e Contratação em Computação"):
                soma+=1.5
            elif(list == "Curricularização da Extensão no Curso de Engenharia de Computação"):
                soma+=1
            elif(list == "ENADE Exame Nacional de Desempenho de Estudantes"):
                soma+=1
            elif(list == "Momento WIE"):
                soma+=2
        if(soma > 20):
            soma = 20

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
    imprimeLista(listapessoas)
    dic = transformar(listapessoas)
    df = pd.DataFrame(dic)
    df.to_excel("Python\ListaCarga.xlsx", index=False)
    

main()


