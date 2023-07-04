import pandas as pd
import ctypes


#Front
import tkinter as tk
def fechar_janela():
    global valor_digitado
    valor = entrada.get()
    # Faça algo com o valor capturado
    valor_digitado = valor
    print("A variável 'valor_digitado' contém:", valor_digitado)

    janela.destroy()

    



# Cria a janela principal
janela = tk.Tk()

janela.title("Digite o nome que o arquivo está")
janela.geometry("300x100")

# Cria a caixa de entrada
entrada = tk.Entry(janela)
entrada.pack()


# Cria o botão para capturar o valor e fechar a janela
botao = tk.Button(janela, text="Salvar Arquivo com esse nome", command=fechar_janela)
botao.pack()

# Inicia o loop principal da janela
janela.mainloop()




def exibir_alerta(titulo, mensagem):
    ctypes.windll.user32.MessageBoxW(0, mensagem, titulo, 0)
# Carrega a planilha original


archive = valor_digitado
try:
    planilha_original = pd.read_excel(f'{archive}.xlsx')
except FileNotFoundError:
     exibir_alerta('Aquivo não encontrado',f'Por gentileza despejar na pasta um arquivo com o nome: {archive}.xlsx ')
    

# Divide a planilha em planilhas menores com 4000 linhas cada uma
num_linhas_por_planilha = 4000
total_linhas = len(planilha_original)

num_planilhas = total_linhas // num_linhas_por_planilha
resto = total_linhas % num_linhas_por_planilha

if resto > 0:
    num_planilhas += 1

for i in range(num_planilhas):
    inicio = i * num_linhas_por_planilha
    fim = inicio + num_linhas_por_planilha
    
    # Extrai as linhas correspondentes a cada planilha
    planilha_dividida = planilha_original.iloc[inicio:fim]
    
    # Salva a planilha dividida em um novo arquivo
    nome_arquivo = f'{valor_digitado}{i+1}.xlsx'
    planilha_dividida.to_excel(nome_arquivo, index=False)

print(f"A planilha original foi dividida em {num_planilhas} planilhas.")
