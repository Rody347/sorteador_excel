# This Python file uses the following encoding: utf-8
import os, sys
import openpyxl

#Criando a planilha
book = openpyxl.Workbook()
#Criando a página
book.create_sheet('Clientes')

#Cadastrando os clientes fictícios
#,telefone,endereco,produto
clientes_page = book['Clientes']
clientes_page.append(['Nome','telefone','endereco','produto participante'])
clientes_page.append(['Paula ','telefone','endereco','produto'])
clientes_page.append(['Fernanda ','telefone','endereco','produto'])
clientes_page.append(['Renan ','telefone','endereco','produto'])
clientes_page.append(['Amanda','telefone','endereco','produto'])
clientes_page.append(['Pedro ','telefone','endereco','produto'])
clientes_page.append(['Paulo ','telefone','endereco','produto'])
clientes_page.append(['Gabriel ','telefone','endereco','produto'])
clientes_page.append(['Guilherme ','telefone','endereco','produto'])
clientes_page.append(['Lucas ','telefone','endereco','produto'])
clientes_page.append(['Ricardo ','telefone','endereco','produto'])
clientes_page.append(['Felipe ','telefone','endereco','produto'])
clientes_page.append(['Filipe ','telefone','endereco','produto'])

book.save('Planilha de Clientes.xlsx')