import tkinter as tk
from tkinter import *
from tkinter import ttk
import sqlite3
import pandas as pd
from openpyxl import Workbook

#Banco de dados utilizado: SQLite
#Linguagem: Python
# criação da tabela no banco:
    # con.execute(''' CREATE TABLE pessoas (
    #         nome text,
    #         idade integer
    #         )
    #     ''')



#Criação de funções e lógica (Model):

#CRUD - Create:
def cadastrarPessoa():
    # conexão com o banco de dados (Repository):
    global con
    con = sqlite3.connect('pessoa.db')

    # responsável pelas ações:
    mensageiro = con.cursor()

    mensageiro.execute(" INSERT INTO pessoas VALUES (:nome, :idade)",
                {
                    'nome':entNome.get(),
                    'idade':entIdade.get()
                }
                )

    #confirmação de criação e fechamento de conexão:
    con.commit()
    con.close()

    #limpa os campos
    entNome.delete(0,"end")
    entIdade.delete(0, "end")


#CRUD - Read:
def read(query): #Lê informações do banco
    con = sqlite3.connect('pessoa.db')
    mensageiro = con.cursor()
    mensageiro.execute(query)
    read = mensageiro.fetchall()
    con.close()

    return read

def listar(): #Função para listar informações do banco
    tree.delete(*tree.get_children())
    select = "SELECT * FROM pessoas"
    linhas = read(select)
    for i in linhas:
        tree.insert("", "end", values=i)

#CRUD - Update:
def selecionarPessoa():
    # limpa os campos
    entNome2.delete(0, "end")
    entIdade2.delete(0, "end")

    #seleciona o item clicado na tabela
    selecionado = tree.focus()
    pessoa_selecionada = tree.item(selecionado, 'values') #Pega nome e idade da pessoa selecionada na tabela

    entNome2.insert(0, pessoa_selecionada[0]) #nome
    entIdade2.insert(0, pessoa_selecionada[1]) #idade


def updatePessoa():
    con = sqlite3.connect('pessoa.db')
    mensageiro = con.cursor()

    selecionado = tree.focus()
    pessoa_selecionada = tree.item(selecionado, 'values')

    #Salvar mudança na tabela:
    tree.item(selecionado, text='', values=(entNome2.get(), entIdade2.get()))

    #Salvar mudança no banco:
    mensageiro.execute('''UPDATE pessoas SET
                nome = :novo
                WHERE idade = :idade''',
                {
                    'novo': entNome2.get(), #Pega o novo nome escrito
                    'idade': pessoa_selecionada[1], #se baseando na idade da pessoa selecionada na tabela
                }

                )

    mensageiro.execute('''UPDATE pessoas SET
                    idade = :nova
                    WHERE idade = :idade''',
                {
                    'nova': entIdade2.get(), #Pega a nova idade escrita
                    'idade': pessoa_selecionada[1], #se baseando na idade da pessoa selecionada na tabela
                }

                )

    con.commit()

    entNome2.delete(0, "end")
    entIdade2.delete(0, "end")

    con.close()


#CRUD - Delete:
def deletarPessoa():
    con = sqlite3.connect('pessoa.db')
    mensageiro = con.cursor()

    selecionado = tree.focus()
    pessoa_selecionada = tree.item(selecionado, 'values')

    mensageiro.execute('''DELETE from pessoas WHERE
                        idade = :idade''',
                {
                    'idade': pessoa_selecionada[1],  #Exclui do banco se baseando na idade da pessoa selecionada na tabela
                }

                )

    con.commit()

    entNome2.delete(0, "end")
    entIdade2.delete(0, "end")

    listar()

    con.close()

def mostrarPessoas(): #Exibe a janela com nome e idade registradas no banco
    con = sqlite3.connect('pessoa.db')

    tabelaInterface = Toplevel()

    global tree
    tree=ttk.Treeview(tabelaInterface, columns=('nome', 'idade'), show='headings')
    tree.column('nome', minwidth=0, width=150)
    tree.column('idade', minwidth=0, width=150)
    tree.heading('nome', text="NOME")
    tree.heading('idade', text="IDADE")
    tree.pack()
    listar()

    global entNome2
    global entIdade2
    # Nome
    lblNome2 = tk.Label(tabelaInterface, text='Nome')
    lblNome2.pack(pady=10)
    entNome2 = tk.Entry(tabelaInterface, text='Nome', width=25)
    entNome2.pack(pady=5)

    # Idade
    lblIdade2 = tk.Label(tabelaInterface, text='Idade')
    lblIdade2.pack(pady=10)
    entIdade2 = tk.Entry(tabelaInterface, text='Idade', width=25)
    entIdade2.pack(pady=5)

    #Botões da tabelaInterface:
    #Selecionar pessoa:
    btSelect = Button(tabelaInterface, text="Selecionar Pessoa", command=selecionarPessoa)
    btSelect.pack(pady=20)
    #Salvar mudanças:
    btUpdate = Button(tabelaInterface, text="Salvar Mudanças", command=updatePessoa)
    btUpdate.pack(pady=10)
    #Deletar pessoa:
    btDelete = Button(tabelaInterface, text="Deletar Pessoa", command=deletarPessoa)
    btDelete.pack(pady=10)

    con.close()


#Exportação para o Excel:
def exportExcel():
    con = sqlite3.connect('pessoa.db')
    mensageiro = con.cursor()

    mensageiro.execute("SELECT *, oid FROM pessoas")
    pessoasCadastradas = mensageiro.fetchall()
    pessoasCadastradas = pd.DataFrame(pessoasCadastradas, columns=['nome', 'idade', 'id'])
    pessoasCadastradas.to_excel('cadastros_pessoas.xlsx')

    con.commit()
    con.close()

    entNome.delete(0, "end")
    entIdade.delete(0, "end")



#criação de interface gráfica (View):
interface = tk.Tk()
interface.title('Cadastro de pessoas')

#labels da interface:
#Nome
lblNome = tk.Label(interface, text = 'Nome')
lblNome.grid(row=0, column=0, padx=10, pady=10)
#Idade
lblIdade = tk.Label(interface, text = 'Idade')
lblIdade.grid(row=1, column=0, padx=10, pady=10)

#entrys da interface:
#Nome
entNome = tk.Entry(interface, text = 'Nome', width=25)
entNome.grid(row=0, column=1, padx=10, pady=10)
#Idade
entIdade = tk.Entry(interface, text = 'Idade', width=25)
entIdade.grid(row=1, column=1, padx=10, pady=10)

#botões da interface:
#Cadastro:
btCadastro = tk.Button(interface, text = 'Cadastrar', command = cadastrarPessoa)
btCadastro.grid(row=2, column=0, padx=10, pady=10, columnspan=2, ipadx=80)
#Exportar para excel:
btExportar = tk.Button(interface, text = 'Exportar para Excel', command = exportExcel)
btExportar.grid(row=3, column=0, padx=10, pady=10, columnspan=2, ipadx=55)
#Mostrar pessoas cadastradas:
btMostrar = tk.Button(interface, text = 'Mostrar Pessoas Cadastradas', command = mostrarPessoas)
btMostrar.grid(row=4, column=0, padx=10, pady=10, columnspan=2, ipadx=30)


interface.mainloop()