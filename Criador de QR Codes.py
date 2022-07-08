# -*- coding: utf-8 -*-
"""Criador de QR Codes"""

# pip install qrcode[pil]
# pip install pandas
# pip install PySimpleGUI

from functools import total_ordering
from tkinter import CENTER
import qrcode
import pandas as pd
import PySimpleGUI as sg

#criar a lista de alunos baseada no excel
df = pd.read_excel("Lista de Alunos.xlsx")
lista_de_nomes = df.values.tolist()

totalcontador = 0
for row in lista_de_nomes:
  totalcontador +=1


#Função para gerar a lista de QR codes com nome e turma do aluno

def gerar_qr_code():
  global contador
  contador = 0
  for row in lista_de_nomes:
    qr = qrcode.QRCode(
      version=1,
      error_correction=qrcode.constants.ERROR_CORRECT_L,
      box_size=10,
      border=4,
    )
    qr.add_data(row)
    qr.make(fit=True)
    contador +=1
    img = qr.make_image(fill_color="black", back_color="white")
    window['progress'].UpdateBar(contador)
    img.save(f"QRCodes Criados/{row}.png")
  
    

"""GUI do aplicativo"""

sg.theme('Material1')

frame_status_progressbar = [
          [sg.ProgressBar(totalcontador, orientation='h', s=(50,10), key='progress')],
          [sg.StatusBar(f'{totalcontador} QR codes na lista.', k='STATUS', justification=CENTER)],
]

frame1 = [
          [sg.Text('1 - Preencher a planilha "Lista de Alunos" com o nome e turma dos alunos da sala/escola')],
          [sg.Text('A "Lista de Alunos também pode ser vir como lista de livros, salas, ou qualquer outra categoria\n que desejar criar QR Codes.')],
          [sg.Text('2 - Clicar no botão "Gerar QR Code"')],
          [sg.Text('3 - Os QRs codes ficarão na pasta "QR codes criados"')],
          [sg.Text('4 - Utilize qualquer aplicativo de leitura de QR codes para fazer a leitura."')],
]

layout = [[sg.Image('logo.png')],
          [sg.Push()],
          [sg.Text('Olá Profissional da Edução.\n')],
          [sg.Push()],
          [sg.Text('Este programa irá automatizar a criação de QR Codes\n para toda sua turma/escola em um único processo')],
          [sg.Push()],
          [sg.Frame('Como usar o programa:', frame1)],
          [sg.Push()],
          [sg.Text('Importante: Não mova a "Lista de Alunos" de Pasta ou o programa não irá funcionar')],
          [sg.Push()],
          [sg.Frame('', frame_status_progressbar)],
          [sg.Push()],
          [sg.Button('Gerar QR Codes'), sg.Button('Sair')]
]

window = sg.Window('Gerador de QR Codes para escolas', layout, element_justification=CENTER, size=(600, 550))


while True:  
    event, values = window.read()
    print(event, values)
    if event == sg.WIN_CLOSED or event == 'Sair':
        break
    if event == 'Gerar QR Codes':
        gerar_qr_code()
        window['STATUS'].Update(f'{contador} QR Codes criados')
        print(contador)
    
window.close()

