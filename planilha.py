# This Python file uses the following encoding: utf-8
import os, sys
import openpyxl

#Criando a planilha
book = openpyxl.Workbook()
#Criando a página
book.create_sheet('Clientes')

#Cadastrando os clientes fictícios
#,telefone,endereco,produto
s_page = book['Clientes']
s_page.append(['Nome','telefone','endereco','produto participante'])
s_page.append(['Paula ','telefone','endereco','produto'])
s_page.append(['Fernanda ','telefone','endereco','produto'])
s_page.append(['Renan ','telefone','endereco','produto'])
s_page.append(['Amanda','telefone','endereco','produto'])
s_page.append(['Pedro ','telefone','endereco','produto'])
s_page.append(['Paulo ','telefone','endereco','produto'])
s_page.append(['Gabriel ','telefone','endereco','produto'])
s_page.append(['Guilherme ','telefone','endereco','produto'])
s_page.append(['Lucas ','telefone','endereco','produto'])
s_page.append(['Ricardo ','telefone','endereco','produto'])
s_page.append(['Felipe ','telefone','endereco','produto'])
s_page.append(['Filipe ','telefone','endereco','produto'])

book.save('Planilha de Clientes.xlsx')