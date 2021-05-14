from openpyxl import Workbook
import os

wb = Workbook()
ws = wb.active


def criar_pagina_planilha():
    pagina = input('Digite o nome da página: ')
    wb.create_sheet(pagina)


def limpar_terminal():
    os.system('cls' if os.name =='nt' else clear)



limpar_terminal()

#Bem vindo
print('Bem vindo!\nPara começar crie uma nova página dentro da planilha.')
criar_pagina_planilha()



while True:
    resposta_criar_planilha = input('Deseja criar mais uma página nesta planilha? (s/n): ')
    if resposta_criar_planilha.lower() == 's':
        criar_pagina_planilha()
    elif resposta_criar_planilha.lower() =='n':
        del wb['Sheet']
        print(wb.sheetnames)
        break



#tratar dados para aceitar maiúscula ou minuscula
planilha_a_manipular = input('Escolha a planilha que deseja manipular: ')
planilha_selecionada = wb[planilha_a_manipular]
lista = ([])

while True:
    lista.append(input('Digite um nome para seu cabeçalho: '))
    resposta_cabecalho = input('Deseja adicionar mais uma coluna? (s/n): ')
    if resposta_cabecalho.lower() == 's':
        pass
    elif resposta_cabecalho.lower() =='n':
        planilha_selecionada.append(lista)
        break

adicionar_mais_dados = input('Deseja adicionair mais dados a essa planilha? (s/n): ')
if adicionar_mais_dados.lower() == 's':
    pagina_escolhida =  input(f'Essas são as planilhas disponíveis > \n{wb.sheetnames}\nEm qual página deseja adicionar mais dados? ')
elif adicionar_mais_dados.lower() == 'n':
    pass

pagina_para_adicionar_dados = wb[pagina_escolhida]
lista = ([])

while True:
    dados_a_adicionar = input('Digite os dados a serem a uma nova linha, separados por vírgula: ')
    lista.append(dados_a_adicionar.split(','))
    for dados in lista:
        pagina_para_adicionar_dados.append(dados)
    resposta_adicionar_nova_linha = input('Adicionar nova linha? (s/n): ')
    if resposta_adicionar_nova_linha.lower() == 's':
        pass
    elif resposta_adicionar_nova_linha.lower() =='n':
        break

nome_para_salvar = input('Qual o nome deseja salvar sua planilha?')

#row - linha
# column - coluna

wb.save(f'{nome_para_salvar}.xlsx')
