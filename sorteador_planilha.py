# This Python file uses the following encoding: utf-8
import os, sys
import openpyxl 
import random

#Especificando a planilha
book = openpyxl.load_workbook('Planilha de Clientes.xlsx')
#Especificando a página
clientes_page = book['Clientes']

#Pedindo a quantidade de elementos que serão sorteados
print('Escolha o número 12(A planilha foi feita com 12 clientes)')
n = input('Insira a quantidade de clientes participantes: ')
#Criando o número sorteador aleatório e limitando ao tamanho do vetor
s = random.randrange(1,n)

#Criando o vetor com os participantes(puxando os dados da planilha) e iterando o sorteio
clientes = []
for col in clientes_page.iter_cols(min_row=2,max_col=1):
    for cell in col:
        clientes.append(cell.value)

#Caso você queira comprovar que o índice e o elemento estão sendo sorteados corretamente, é só descomentar essas duas linhas e rodar o código
# print(clientes)
# print(s)

#Imprimindo o cliente sorteado
print(clientes[s])