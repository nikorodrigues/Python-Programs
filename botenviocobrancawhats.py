"""
docstring é definida por 3 aspas Duplas("""""")
"""

import webbrowser
from urllib.parse import quote
from time import sleep
import os
import openpyxl
import pyautogui

# Abrir WhatsApp Web no navegador -> import bib webbrowser
webbrowser.open('https://web.whatsapp.com')
# Dar tempo para tarefas executarem -> importar sleep -> from time import sleep
sleep(5)

# Carregar planilha em excel
# Indicar qual página da planilha deve ser lida
workbook = openpyxl.load_workbook('clientes2.xlsx')
pagina = workbook['contatos']

# Criar variaveis com os dados da planilha (nome, telefone, vencimento)
# Usar o comando For linha in para ler cada linha da planilha
for linha in pagina.iter_rows(min_row=2):
    # Extrair as informações que eu preciso para rodar a automação
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    # Criar uma mensagem personalizada -> Utilizando uma fstring
    # Importar datetime, já que vencimentos
    # geralmente tem o formato dd/mm/aaaa, não acompanhado de horas
    mensagem = f'Olá {nome}, Seu boleto que vence em *{vencimento.strftime(
        "%d/%m/%Y")}* está disponível para pagamento, clique no link https://estesitenaoexiste.com.br/ para pagar. *ISTO É APENAS UM TESTE, IGNORE A MENSAGEM.*'

    # Criar um link editável para enviar para telefones diferentes com textos diferentes (API DO WHATSAPP WEB) -> https://web.whatsapp.com/send?phone=5511999999999&text=safasdasf -> Utilizando uma fstring formatada(editável)
   # from urllib.parse import quote(que permite editar textos em links)
    try:

        link = f'https://web.whatsapp.com/send?phone={
            telefone}&text={quote(mensagem)}'

        webbrowser.open(link)
        sleep(5)
        pyautogui.press('enter')
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
        fechar = pyautogui.locateCenterOnScreen('usar_nesta_janela.png')
        pyautogui.click(fechar[0], fechar[1])
        # Clicar em Usar nesta janela para retornar à tela principal do whatsapp
    except Exception as error:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
# Criar um arquivo excel chamado Relatório com data, informando"
# cobrança enviada com sucesso" ou "envio de cobrança falhou"
