from curses import window
from PySimpleGUI import Window, Text, ProgressBar, Button

# def cancelar():
#   driver.quit()



layout = [
  [Text('Aguarde enquanto o programa envia as mensagens',background_color='#0e0e24')],
  [ProgressBar(1000,orientation='h', size=(50,20), border_width=4, key='-PB-', bar_color=('#45E28D', 'white'))],
  [Button('Rodar', key='-RODAR-'),Button('Cancelar', key='-CANCELAR-')]
]


window_pb = Window('Enviando Mensagens', layout, icon='SEIBOT.ico', background_color='#0e0e24')


window_pb.close()
