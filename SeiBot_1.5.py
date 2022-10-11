from ast import Continue
from sqlite3 import dbapi2
from tkinter import filedialog
import tkinter
from turtle import width
#from numpy import pad
import pyautogui
from tkinter import *
from tkinter import ttk
from lib2to3.pgen2.token import OP
import webbrowser
from os import times_result
from random import expovariate
import pandas as pd
from pandas import options
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.edge.service import Service as EdgeService # Usar no Edge
from webdriver_manager.microsoft import EdgeChromiumDriverManager # Usar no Edge
from selenium.webdriver.edge.options import Options # Usar no Edge
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
import selenium.webdriver.support.expected_conditions as ec
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from tkinter import Label, Tk
from tkinter.filedialog import askopenfilename, askopenfilenames
import sys
import os
from time import sleep
import urllib
from myimages import *
import datetime as dt
import threading


# Getting actual date of the System
datahoje = str(dt.date.today())

# Setting an experation date of the program
if datahoje < ('2023-04-02'):

  #Creates a alert prompt
  alerta = pyautogui.alert(text="""Esse programa foi testado com lista de 300 contatos.
O WhatsApp não gosta de automação, portanto, use com sabedoria. O robô simula a interação humana, por isso demora
um pouco. Me isento de quaisquer problemas que possa ter pelo mal uso da ferramenta!""", title='AVISO')

  # Gets computer name
  nome_computador = os.getlogin()

  # Creating app window with Tkinter
  app=Tk()
  app.title("SEIBot")
  app.geometry("700x600")
  app.minsize(700,600)
  app.configure(background="#0e0e24")
  iconimg = PhotoImage(file='logoapp.png')
  app.iconphoto(False, iconimg)

# ______ Functions that centralize the app window. ______#
  def center(win):
      # :param win: the main window or Toplevel window to center

      # Apparently a common hack to get the window size. Temporarily hide the
      # window to avoid update_idletasks() drawing the window in the wrong
      # position.
      win.update_idletasks()  # Update "requested size" from geometry manager

      # define window dimensions width and height
      width = win.winfo_width()
      frm_width = win.winfo_rootx() - win.winfo_x()
      win_width = width + 2 * frm_width

      height = win.winfo_height()
      titlebar_height = win.winfo_rooty() - win.winfo_y()
      win_height = height + titlebar_height + frm_width

      # Get the window position from the top dynamically as well as position from left or right as follows
      x = win.winfo_screenwidth() // 2 - win_width // 2
      y = win.winfo_screenheight() // 2 - win_height // 2

      # this is the line that will center your window
      win.geometry('{}x{}+{}+{}'.format(width, height, x, y))

      # This seems to draw the window frame immediately, so only call deiconify()
      # after setting correct window position
      win.deiconify()

  app.attributes('-alpha', 0.0)
  center(app)
  app.attributes('-alpha', 1.0)
  #---------------------------------------------------------------------------------#

  # Variable that saves which browser the user is using
  browser = pyautogui.confirm(text='QUAL NAVEGADOR VOCÊ USA COMO PADRÃO?', title='NAVEGADOR PADRÃO', buttons=['CHROME', 'OUTRO'])

 #______ Function that allows the user to load a file ______#
  def Carregar_Arq():
    abrir = askopenfilename()
    global dx
    dt = abrir
    dx = pd.read_excel(dt)
    # setting grid and style
    arqcarregado = Label(app,text="Arquivo Carregado",background='#0e0e24',foreground='#45E28D')
    arqcarregado.grid(row=2, column=2)
    app.grid_rowconfigure(2, weight=1)
    return dt

# ____________ Code to Run on Chrome ____________#
  # Function to keep the program to crash while in use
  def tt_Chrome():
    t1=threading.Thread(target=Enviar_Msg_Chrome)
    t1.start()

  #______ Function used to send the messages ______#
  def Enviar_Msg_Chrome():
    #Defining grid and style of the progress bar
    barra_progresso = ttk.Progressbar(app, orient=HORIZONTAL, length=250, mode='determinate')
    barra_progresso.grid(row=6, column=2)
    app.grid_rowconfigure(6, weight=1)

    df = dx
    mensagem = msg.get("1.0","end-1c") # This code allows to get the whole text in the box.
      # Abrir navegador em Cache
    options = Options()
    #open defaut profile of the user's browser
    options.add_argument(f"user-data-dir=C:/Users/{nome_computador}/AppData/Local/Google/Chrome/User Data")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver,100)
    site = ('https://web.whatsapp.com')
    driver.get(site)
    semwhats = [] #Empty list to store the invalid numbers
    count = 0
    
    # Waits WhatsApp Web to load
    wait.until(ec.text_to_be_present_in_element((By.XPATH, '//*[@id="app"]/div/div/div[4]/div/div/div[2]/div[1]/h1'),'WhatsApp Web'))

    # Prompt window to run the program
    rodar = pyautogui.confirm(text='Clique em Ok para começar a enviar as mensagens', title='Começar Programa')

    if rodar == 'OK':
      # Conta quantas linhas tem no DataFrame
      contar_linhas = df[df.columns[0]].count()
      #Acessar a base de dados e colocar no código pra ser enviado)
      porcentagem = Label(app, textvariable=percent, background='#0e0e24',foreground='#fff')
      porcentagem.grid(row=7, column=2)
      percent.set(str('0'+'%'))
      app.grid_rowconfigure(7, weight=1)
      
      # goes throug the list uploaded by user and gets name and number
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

          # variable that divides by 100 the number of rows of the table
          prg = 100 / contar_linhas
          # adds gradually the value of the variable above to the Progress Bar
          x = barra_progresso['value'] = barra_progresso['value'] + prg
          xx = int(x)
          percent.set(str(xx)+"%")
          app.update_idletasks()

        # Setting an exception to the code above
        except Exception:
          # it adds the value to the list created above
          semwhats.append(nome + " " + str(numero))
          count = count + 1
          prg = 100 / contar_linhas
          x = barra_progresso['value'] = barra_progresso['value'] + prg # Adiciona gradualmente o valor da divsão da variável acima para a barra de
          xx = int(x)
          percent.set(str(xx)+"%")
          app.update_idletasks()
          pass

      # Condition to save the invalids results
      if count > 0:
        invalidos = pyautogui.confirm(text='Deseja ver os números inválidos?', title='Números Inválidos')
        if invalidos =='OK':
          file = filedialog.asksaveasfile(defaultextension='.xlsx', filetypes=[('Arquivo TXT','.txt'),('Excel', '.xlsx'),('Todos','.*')])
          file.write(str(semwhats))
          file.close()
      else:
        pass

      #Prompt to asks if the user wishes to close the browser
      pyautogui.alert(text='Terminamos por aqui. Pode fechar o navegador!', title='TERMINEI!')
      driver.quit()
    else:
      Continue
# ____________ END OF CODE TO RUN IN CHROME ____________#


# ____________ CODE TO RUN ON EDGE ____________#

   # Essa função, permite com que meu programa não trava enquanto envia as mensagens
  def tt_edge():
    t1=threading.Thread(target=Enviar_Msg_Edge)
    t1.start()

  def Enviar_Msg_Edge():
    barra_progresso = ttk.Progressbar(app, orient=HORIZONTAL, length=250, mode='determinate')
    barra_progresso.grid(row=6, column=2)
    app.grid_rowconfigure(6, weight=1)

    df = dx
    mensagem = msg.get("1.0","end-1c")
      # Abrir navegador em cache
    edge_options = Options()
    edge_options.add_argument(f"user-data-dir=C:\\Users\\{nome_computador}\\AppData\\Local\\Microsoft\\Edge\\User Data")


    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=edge_options)
    wait = WebDriverWait(driver,100)
    site = ('https://web.whatsapp.com')
    driver.get(site)
    semwhats = []
    count = 0

    wait.until(ec.text_to_be_present_in_element((By.XPATH, '//*[@id="app"]/div/div/div[4]/div/div/div[2]/div[1]/h1'),'WhatsApp Web'))

    rodar = pyautogui.confirm(text='Clique em Ok para começar a enviar as mensagens', title='Começar Programa')

    if rodar == 'OK':

      #Acessar a base de dados e colocar no código pra ser enviado
      # Conta quantas linhas tem no DataFrame
      contar_linhas = df[df.columns[0]].count()
      #Acessar a base de dados e colocar no código pra ser enviado)
      porcentagem = Label(app, textvariable=percent, background='#0e0e24',foreground='#fff')
      porcentagem.grid(row=7, column=2)
      percent.set(str('0'+'%'))
      app.grid_rowconfigure(7, weight=1)
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
          # Variável que divide por 100 a quantidade de linhas que tem na tabelma
          prg = 100 / contar_linhas
          x = barra_progresso['value'] = barra_progresso['value'] + prg # Adiciona gradualmente o valor da divsão da variável acima para a barra de
          xx = int(x)
          percent.set(str(xx)+"%")
          app.update_idletasks()
        except Exception:
          semwhats.append(nome + " " + str(numero))
          count = count + 1
          # Teste de PrgressBar
          prg = 100 / contar_linhas
          x = barra_progresso['value'] = barra_progresso['value'] + prg # Adiciona gradualmente o valor da divsão da variável acima para a barra de
          xx = int(x)
          percent.set(str(xx)+"%")
          app.update_idletasks()
          pass

      if count > 0:
        invalidos = pyautogui.confirm(text='Deseja ver os números inválidos?', title='Números Inválidos')
        if invalidos =='OK':
          file = filedialog.asksaveasfile(defaultextension='.txt', filetypes=[('Arquivo TXT','.txt'),('Todos','.*')])
          file.write(str(semwhats))
          file.close()
      else:
        pass

      pyautogui.alert(text='Terminamos por aqui. Pode fechar o navegador!', title='TERMINEI!')
      driver.quit()
    else:
      Continue
# ____________ END OF CODE TO RUN ON EDGE ____________#

  pic = logostring
  percent = StringVar()
  render = PhotoImage(data=pic)
  logo = Label(app, image=render)

#__________ GRIDING | LAYOUT _________#
  
  carregar_arq = Button(app, text='Carregar Arquivo', command=Carregar_Arq)
  txbox = Label(app,text='Digite a mensagem que deseja enviar para todos',background='#0e0e24',foreground='#fff')
  msg = Text(app,width=50, height=10)
  msg.insert(END,"Digite aqui sua mensagem.")
  if browser == 'CHROME':
    btn_enviar = Button(app, text="Enviar", command=tt_edge)
  if browser == 'OUTRO':
    btn_enviar = Button(app, text="Enviar", command=tt_Chrome)
  info = Label(app,text="SeiBot. Versão 1.5 - Desenvolvido por Winston de Moura",background='#0e0e24',foreground='#fff')

  logo.grid(row=0, column=2,padx=15, pady=15)
  carregar_arq.grid(row=1, column=2, pady=(0,10))
  txbox.grid(row=3, column=2)
  msg.grid(row=4, column=2)
  btn_enviar.grid(row=5, column=2, pady=15)
  info.grid(row=8, column=2)

  # Griding
  app.grid_rowconfigure(0, weight=3) # Logo
  app.grid_rowconfigure(1, weight=2) # Botão Carregar
  app.grid_rowconfigure(3, weight=2) # Mensagem da Caixa de Texto
  app.grid_rowconfigure(4, weight=1) # Caixa de Texto
  app.grid_rowconfigure(5, weight=3) # Botão Enviar
  app.grid_rowconfigure(8, weight=1) # Info

  app.grid_columnconfigure(2, weight=1)


  app.mainloop()

else:
  pyautogui.alert(text="Programa Expirou. Baixe-o novamente do site: <href:http://seibot.ml>", title='Programa Expirou')