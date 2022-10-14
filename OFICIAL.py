from curses import window
from turtle import color, right
from webbrowser import BackgroundBrowser
from PySimpleGUI import Window, Button, Image, Text, Column, theme, MLine,Push, HSeparator, read_all_windows, WIN_CLOSED
import pandas as pd
from pandas import options
import sys
import os
import pyautogui
import urllib
import time

from tkinter.filedialog import askopenfilename, askopenfilenames
import threading

nome_computador = os.getlogin()
# ________________________________Janelas do programa e estilos___________________________#

def janela_ini():
  theme('DarkGray 14')

  layout1=[[Image(filename='assets\logoapp.png', background_color='#0e0e24', size=(100,100))]]
  layout2=[[Text('Qual seu navegador padrão?', background_color='#0e0e24')]]
  layout3=[[Push(background_color='#0e0e24'),Button('Chrome', key='-CHROME-'), Button('Outro', key='-OUTRO-'),Push(background_color='#0e0e24')]]
  layout4=[[Text('', background_color='#0e0e24')]]
  layout5=[[MLine(default_text='Esse programa foi testado com lista de 300 contatos. O WhatsApp não gosta de automação, portanto, use com sabedoria. O robô simula a interação humana, por isso demora um pouco. Me isento de quaisquer problemas que possa ter pelo mau uso da ferramenta!', size=(70,5)), Push()]]
  layout6=[[Push(),Text('Desenvolvido por Winston de Moura - Versão 1.5',background_color='#0e0e24'), Push()]]

  layout = [
    [Column(layout1,background_color='#0e0e24')],
    [Column(layout2,background_color='#0e0e24')],
    [Column(layout3,background_color='#0e0e24')],
    [Column(layout4,background_color='#0e0e24')],
    [Column(layout5,background_color='#0e0e24')],
    [Column(layout6,background_color='#0e0e24')]
    ]

  return Window(
  'SeiBot - Versão 1.5',
  icon='assets\SEIBOT.ico',
  background_color='#0e0e24',
  size=(600, 400),
  layout=layout,
  element_justification='c',
  finalize=True
)

def janela_chrome():

  layout1=[[Image(filename='assets\logoapp.png',background_color='#0e0e24', size=(100,70))]]
  #layout2=[[FileBrowse(button_text='Carregar Arquivo', key='-LOAD-')]]
  layout2=[[Button('Carregar Arquivo', key='-LOAD-')]]
  layout7=[[Text('', text_color='#45E28D', key='-ARQCARREGADO-',background_color='#0e0e24')]]
  layout3=[[Text('Digite a mensagem que deseja enviar para todos',background_color='#0e0e24')]]
  layout4=[[MLine(default_text='Digite aqui sua mensagem.', size=(50,12), key='-TEXTO-')]]
  layout5=[[Button('Enviar', key='-ENVIAR-')]]
  layout6=[[Text('SeiBot - Versão 1.5 - Desenvolvido por Winston de Moura',background_color='#0e0e24')]]
  layout8=[[]]

  layout = [
    [Column(layout1,background_color='#0e0e24')],
    [Column(layout8,background_color='#0e0e24')],
    [Column(layout2,background_color='#0e0e24')],
    [Column(layout7,background_color='#0e0e24')],
    [Column(layout3,background_color='#0e0e24')],
    [Column(layout4,background_color='#0e0e24')],
    [Column(layout5,background_color='#0e0e24')],
    [Column(layout6,background_color='#0e0e24')],
    ]

  return Window(
    'SeiBot - Chrome',
    icon='assets\SEIBOT.ico',
    background_color='#0e0e24',
    size=(600,500),
    layout=layout,
    element_justification='c',
    finalize=True
  )

def janela_edge():
  layout1=[[Image(filename='assets\logoapp.png',background_color='#0e0e24', size=(100,70))]]
  #layout2=[[FileBrowse(button_text='Carregar Arquivo', key='-LOAD-')]]
  layout2=[[Button('Carregar Arquivo', key='-LOAD-')]]
  layout7=[[Text('', text_color='#45E28D', key='-ARQCARREGADO-',background_color='#0e0e24')]]
  layout3=[[Text('Digite a mensagem que deseja enviar para todos',background_color='#0e0e24')]]
  layout4=[[MLine(default_text='Digite aqui sua mensagem.', size=(50,12), key='-TEXTO-')]]
  layout5=[[Button('Enviar', key='-ENVIAR-')]]
  layout6=[[Text('SeiBot - Versão 1.5 - Desenvolvido por Winston de Moura',background_color='#0e0e24')]]
  layout8=[[]]

  layout = [
    [Column(layout1,background_color='#0e0e24')],
    [Column(layout8,background_color='#0e0e24')],
    [Column(layout2,background_color='#0e0e24')],
    [Column(layout7,background_color='#0e0e24')],
    [Column(layout3,background_color='#0e0e24')],
    [Column(layout4,background_color='#0e0e24')],
    [Column(layout5,background_color='#0e0e24')],
    [Column(layout6,background_color='#0e0e24')],
    ]

  return Window(
    'SeiBot - Edge',
    icon='assets\SEIBOT.ico',
    background_color='#0e0e24',
    size=(600,500),
    layout=layout,
    element_justification='c',
    finalize=True
  )


#___________________________ FUNÇÕES DRIVE ______________________________#

def tt_Chrome():
  t1=threading.Thread(target=Enviar_Msg_Chrome)
  t1.start()

def tt_Edge():
  t1=threading.Thread(target=Enviar_Msg_Edge)
  t1.start()

def carregar_arquivo():
  global arq
  arq =  askopenfilename()
  return arq

def Enviar_Msg_Chrome():

  #__________________Chrome_WebDriver ________________#
  from selenium import webdriver
  from selenium.webdriver.chrome.options import Options
  from selenium.webdriver.support.ui import WebDriverWait
  from selenium.webdriver.common.by import By
  from selenium.webdriver.common.keys import Keys
  from selenium.webdriver.support.wait import WebDriverWait
  import selenium.webdriver.support.expected_conditions as ec
  from selenium import webdriver
  from selenium.webdriver.chrome.service import Service
  from webdriver_manager.chrome import ChromeDriverManager
  #__________________________________________________________________#


  db = arq
  df = pd.read_excel(db)
  mensagem = values['-TEXTO-']
    # Abrir navegador em Cache
  chrome_options = Options()
  chrome_options.add_argument(f"user-data-dir=C:/Users/{nome_computador}/AppData/Local/Google/Chrome/User Data")

  driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
  wait = WebDriverWait(driver,100)
  site = ('https://web.whatsapp.com')
  driver.get(site)
  semwhats = []
  count = 0

  # while len(driver.find_elements(By.ID,"side")) < 1:
  #   time.sleep(1)
  # Esperar o WhatsApp Web carregar
  wait.until(ec.text_to_be_present_in_element((By.XPATH, '//*[@id="app"]/div/div/div[4]/div/div/div[2]/div[1]/h1'),'WhatsApp Web'))

  rodar = pyautogui.confirm(text='Clique em Ok para começar a enviar as mensagens', title='Começar Programa')

  if rodar == 'OK':
    # Conta quantas linhas tem no DataFrame
    contar_linhas = df[df.columns[0]].count()
    #Acessar a base de dados e colocar no código pra ser enviado)
    for i, nome in enumerate(df['Nome']):
      try:
        numero = str(df.loc[i, "Número"])
        texto = urllib.parse.quote(f"{mensagem}")
        if numero[0:2] == ('55'):
          link =  f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
        else:
          link =  f"https://web.whatsapp.com/send?phone={'55'+numero}&text={texto}"
        driver.get(link)
        time.sleep(10)
        driver.find_element(By.XPATH,"//span[@data-icon='send']").click()
        time.sleep(6)

      except Exception:
        semwhats.append(nome + " " + str(numero))
        count = count + 1
        pass

    if count > 0:
      invalidos = pyautogui.confirm(text='Deseja ver os números inválidos?', title='Números Inválidos')
      if invalidos =='OK':
        from tkinter import filedialog
        file = filedialog.asksaveasfile(defaultextension='.xlsx', filetypes=[('Arquivo TXT','.txt'),('Excel', '.xlsx'),('Todos','.*')])
        file.write(str(semwhats))
        file.close()
    else:
      pass

    pyautogui.alert(text='Terminamos por aqui. Pode fechar o navegador!', title='TERMINEI!')
    driver.quit()

def Enviar_Msg_Edge():
  #__________Driver_Edge__________#
  from selenium import webdriver
  from selenium.webdriver.edge.service import Service as EdgeService # Usar no Edge
  from webdriver_manager.microsoft import EdgeChromiumDriverManager # Usar no Edge
  from selenium.webdriver.edge.options import Options # Usar no Edge
  from selenium.webdriver.support.ui import WebDriverWait
  from selenium.webdriver.common.by import By
  from selenium.webdriver.common.keys import Keys
  from selenium.webdriver.support.wait import WebDriverWait
  import selenium.webdriver.support.expected_conditions as ec
  #______________________________#

  db = arq
  df = pd.read_excel(db)
  mensagem = values['-TEXTO-']
    # Abrir navegador em cache
  edge_options = Options()
  #edge_options.binary_location = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
  edge_options.add_argument(f"user-data-dir=C:\\Users\\{nome_computador}\\AppData\\Local\\Microsoft\\Edge\\User Data")


  driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=edge_options)
  wait = WebDriverWait(driver,100)
  site = ('https://web.whatsapp.com')
  driver.get(site)
  semwhats = []
  count = 0

  # Esperar o WhatsApp Web carregar
  wait.until(ec.text_to_be_present_in_element((By.XPATH, '//*[@id="app"]/div/div/div[4]/div/div/div[2]/div[1]/h1'),'WhatsApp Web'))

  rodar = pyautogui.confirm(text='Clique em Ok para começar a enviar as mensagens', title='Começar Programa')

  if rodar == 'OK':

    #Acessar a base de dados e colocar no código pra ser enviado
    # Conta quantas linhas tem no DataFrame
    contar_linhas = df[df.columns[0]].count()
    #Acessar a base de dados e colocar no código pra ser enviado)
    for i, nome in enumerate(df['Nome']):
      try:
        numero = str(df.loc[i, "Número"])
        texto = urllib.parse.quote(f"{mensagem}")
        if numero[0:2] == ('55'):
          link =  f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
        else:
          link =  f"https://web.whatsapp.com/send?phone={'55'+numero}&text={texto}"
        driver.get(link)
        time.sleep(10)
        driver.find_element(By.XPATH,"//span[@data-icon='send']").click()
        time.sleep(6)

      except Exception:
        semwhats.append(nome + " " + str(numero))
        count = count + 1
        pass

    if count > 0:
      invalidos = pyautogui.confirm(text='Deseja ver os números inválidos?', title='Números Inválidos')
      if invalidos =='OK':
        from tkinter import filedialog
        file = filedialog.asksaveasfile(defaultextension='.txt', filetypes=[('Arquivo TXT','.txt'),('Todos','.*')])
        file.write(str(semwhats))
        file.close()
    else:
      pass

    pyautogui.alert(text='Terminamos por aqui. Pode fechar o navegador!', title='TERMINEI!')
    driver.quit()

def arq_carregado():
  linha = [[Text('Arquivo Carregado')]]


# Janelas Iniciais
janela1, janela2, janela3 = janela_ini(), None, None
# Loop para leitura dos eventos
while True:
  window,event,values = read_all_windows()
  if window == janela1 and event == WIN_CLOSED:
    break
  if window == janela1 and event == '-OUTRO-':
    janela2 = janela_chrome()
    janela1.hide()
  if window == janela2 and event == '-LOAD-':
    carregar_arquivo()
    window['-ARQCARREGADO-'].update('Arquivo Carregado')
  if window == janela2 and event == '-ENVIAR-':
    tt_Chrome()
  if window == janela2 and event == WIN_CLOSED:
    janela2.hide()
    janela1.un_hide()
  if window == janela1 and event == '-CHROME-':
    janela3 = janela_edge()
    janela1.hide()
  if window == janela3 and event == '-LOAD-':
    carregar_arquivo()
    window['-ARQCARREGADO-'].update('Arquivo Carregado')
  if window == janela3 and event == '-ENVIAR-':
    tt_Edge()
  if window == janela3 and event == WIN_CLOSED:
    janela3.hide()
    janela1.un_hide()

window.close()

# Lógica dos botões
