"""
Oi mãe/pai, não se assustem, eu só estou testando um recurso de programação.
Não, o seu telefone não foi hackeado, é apenas uma mensagem automatizada por um bot que eu estou praticando.
"""
# Descrever os passos manuais e depois escrever os códigos
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(60)


#ler planilha e guardar informações sobre, nome, telefone e vencimento
workbook = openpyxl.load_workbook('teste_bot.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
    #Estrutura da planilha: nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha [1].value
    #vencimento = linha [2].value
    #print(nome , telefone) para ver se está achando corretamente
    
    mensagem = f'Oi {nome}, não se assuste, eu só estou testando um recurso de programação.'
    
    #Criar links personalizados do whats e enviar mensagens para cada cliente
    # https://web.whatsapp.com/send?phone=5555555&text="aaaaaaaaaa"
    
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'  
        webbrowser.open(link_mensagem_whatsapp)
        sleep(30)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(10)
        pyautogui.click(seta[0],seta[1])
        sleep(10)
        pyautogui.hotkey('ctrl','w')
        sleep(10)
    except:
        print(f'Não deu certo no contato {nome}')
        with open('erros.csv','a',newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')



