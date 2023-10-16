import os 
import time
import openpyxl as excel
def arquivo_existe(nome):
	try:
		arquivo = open(nome,"rt")
		arquivo.close()
	except FileNotFoundError:
		return False
	else:	
		return True

def criar_arquivo(nome):
	try:
		arquivo = open(nome,"wt")
		arquivo.close()
	except:
		print("Arquivo nao criado")
	else: 
		print("Arquivo criado")


def cadastrar(nome):
	nome_produto = str(input("Digite o nome do produto: "))
	preco_produto = str(input("Digite o preco do produto: "))
	str_final = nome_produto + ";" + preco_produto + "\n"
	arquivo = open(nome,"at")
	arquivo.write(str_final)
	arquivo.close()

def consultar(nome):
	arquivo = open(nome,"rt")
	for linha in arquivo:
		dado = linha.split(";")
		dado[1] = dado[1].replace("\n","")
		print("Nome: ",dado[0])
		print("Preco: R$",dado[1])
	time.sleep(5)
	#os.system("cls")

def consultar_um(nome):
	nome_produto = str(input("Digite o nome do produto que deseja consultar: "))
	arquivo = open(nome,"rt")
	for linha in arquivo:
		dado = linha.split(";")
		dado[1] = dado[1].replace("\n","")
		if nome_produto == dado[0]:
			print("Nome: ",dado[0])
			print("Preco: R$",dado[1])
	time.sleep(5)
	#os.system("cls")
	

def alterar(nome):
	lista = []
	arquivo = open(nome,"rt")
	for linha in arquivo:
		dado = linha.split(";")
		lista.append(dado[0])
		lista.append(dado[1])
	arquivo.close()
	nome_produto = str(input("Deseja alterar qual produto? "))
	for i in range(0,len(lista),2):
		if lista[i] == nome_produto:
			lista[i] = str(input("Digite o novo nome do produto: "))
			novo_preco = str(input("Digite o novo preco do produto: "))
			novo_preco = novo_preco + "\n"
			lista[i+1] = novo_preco

	arquivo = open(nome,"wt")
	for i in range(0,len(lista)):
		if i % 2 == 0:
			nome = lista[i] + ";"
			string_final = nome + lista[i+1]
			arquivo.write(string_final)
	arquivo.close()
	#os.system("cls")

def excluir(nome):
	lista = []
	nome_produto = str(input("Digite o nome do produto que deseja excluir: "))

	arquivo = open(nome,"rt")
	for linha in arquivo:
		dado = linha.split(";")
		lista.append(dado[0])
		lista.append(dado[1])
	arquivo.close() 

	for i in range(0,len(lista)):
		if lista[i] == nome_produto:
			nome_excluir = lista[i]
			preco_excluir = lista[i+1]
			break
	lista.remove(nome_excluir)
	lista.remove(preco_excluir)
	
	arquivo = open(nome,"wt")
	for i in range(0,len(lista)):
		if i % 2 == 0:
			nome = lista[i] + ";"
			string_final = nome + lista[i+1]
			arquivo.write(string_final)
	arquivo.close()
	#os.system("cls")
	
def salvar_excel(nome):
	lista = []
	livro = excel.Workbook()
	livro_produtos = livro['Sheet']
	arquivo = open(nome,"rt")
	for linha in arquivo:
		lista = []
		dado = linha.split(";")
		lista.append(dado[0])
		lista.append(dado[1])
		livro_produtos.append(lista)
	arquivo.close()
	livro.save("produtos.xlsx")


def menu():
	op = None
	nome = "produtos.txt"
	if arquivo_existe(nome) == False:
		criar_arquivo("produtos.txt")

	while op != 0:
		print("[1] - Cadastrar produto")
		print("[2] - Consultar todos")
		print("[3] - Consultar um")
		print("[4] - Alterar produto")
		print("[5] - Excluir produto")
		print("[6] - Salvar no excel")
		print("[0] - Sair")
		op = int(input(""))
		if op == 0: 
			os.system("cls")
			print("Saindo...")
		elif op == 1:
			os.system("cls")
			cadastrar(nome)
		elif op == 2:
			os.system("cls")
			consultar(nome)
		elif op == 3:
			os.system("cls")
			consultar_um(nome)
		elif op == 4:
			os.system("cls")
			alterar(nome)
		elif op == 5:
			os.system("cls")
			excluir(nome)
		elif op == 6:
			os.system("cls")
			salvar_excel(nome)
		else:
			print("Opcao invalida")
menu()