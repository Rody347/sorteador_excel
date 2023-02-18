# This Python file uses the following encoding: utf-8
import os, sys
import openpyxl 
import random

book = openpyxl.load_workbook('Planilha de Clientes.xlsx')
clientes_page = book['Clientes']
n = input('Insira a quantidade de clientes participantes: ')
s = random.randrange(1,n)
clientes = []

for col in clientes_page.iter_cols(min_row=2,max_col=1):
    for cell in col:
        clientes.append(cell.value)
print(clientes[s])            




# n = input('Insira a quantidade de clientes participantes: ')
# #Gerando número aleatório
# s = random.randrange(1,n)

# def sorteador():
#     s = random.randrange(1,12)
    
# def getting_clientes():
#     clientes = []
#     for col in clientes_page.iter_cols(min_row=2,max_col=1):
#         for cell in col:
#             clientes.append(cell.value)
