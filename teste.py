import time
from tkinter import *
#from winotify import Notification
import mouse
import pyautogui
from PySimpleGUI import PySimpleGUI as sg

#layout da janela
sg.theme('Reddit')
layout = [
	[sg.Text('Senha Atual: motor6046***')],
    [sg.Text('Insira a Senha!'),sg.Input(key = 'senha',size=(20,1))],
	[sg.Button('Entrar')]
]

#janela
janela = sg.Window('Tela de Senha', layout)
eventos, valores = janela.read()
if eventos == sg.WINDOW_CLOSED:
	exit()
if eventos == 'Entrar':
	if not valores['senha']:
		exit()
	else:
		senha = valores['senha']
		print(senha)
    
#planilha com os dados
#planilha = load_workbook ('D:\Automação Preço LG\AmbienteVirtual\Planilha de preco.xlsx')
#aba_ativa = planilha['PREÇO'] # aba com os preços

#login e senha
usuario = 'BR002199'
password = senha

#iniciando o navegador
#service = Service(ChromeDriverManager().install())
#nav = webdriver.Chrome(service=service)
#nav.maximize_window()

#link do site
#nav.get('https://saclge.com.br/svcportal/oow/requests/create')

#login no site
#nav.find_element('id', 'username').send_keys(usuario)
#nav.find_element('id','password').send_keys(password)
#nav.find_element('xpath','//*[@id="landing-choice"]/div/div/form/div/div[2]/div[1]/center/button').click()

arq_img1 = pyautogui.locateOnScreen('D:\Automação Preço LG\AmbienteVirtual\Cod_peça.PNG')

if arq_img1 == None:
	layout2 = [	[sg.Text('SENHA INCORRETA')] ]
	janela2 = sg.Window('AVISO', layout2)
	evento = janela2.read()
	if evento == sg.WINDOW_CLOSED:
		exit()
	else:
		exit()

print('programa em acão')
#arq_img1 = pyautogui.locateOnScreen('D:\Automação Preço LG\AmbienteVirtual\Cod_peça.PNG')
#print(arq_img1)

# layout = [
# 	[sg.Text('Senha Atual: motor6046***')],
#     [sg.Text('Insira a Senha!'),sg.Input(key = 'senha',size=(20,1))],
# 	[sg.Button('Entrar')]
# ]

# #janela
# janela = sg.Window('Tela de Senha', layout)

# eventos, valores = janela.read()
# #valores['senha'] = 1
# if eventos == sg.WINDOW_CLOSED:
# 	exit()
# if eventos == 'Entrar':
# 	if not valores['senha']:
# 		exit()
# 	else:
# 		senha = valores['senha']
# 		print(senha)
		

# print('programa em acão')



# #ler evento
# while True:
# 	eventos, valores = janela1.read()
# 	if eventos == sg.WINDOW_CLOSED:
# 		exit()
# 	if eventos == 'Entrar':
# 		if valores['senha'] != None:
# 			senha = valores['senha']
# 			print(senha)
# 			break






# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager

# usuario = 'BR002199'
# password = 'motor6046***' #input()

# service = Service(ChromeDriverManager().install())
# nav = webdriver.Chrome(service=service)
# nav.maximize_window()
# nav.get('https://saclge.com.br/svcportal/oow/requests/create')


# nav.find_element('id', 'username').send_keys(usuario)
# #pyautogui.typewrite(str(usuario))

# nav.find_element('id','password').send_keys(password)
# #pyautogui.write(str(password))

# nav.find_element('xpath','//*[@id="landing-choice"]/div/div/form/div/div[2]/div[1]/center/button').click()



# time.sleep(5)














# while True:
#     print(mouse.get_position())
#     time.sleep(2)



# avs = Tk()
# avs.title("Aviso")
# avs.geometry("300x100")

# lb_aviso = Label(avs,text="Favor aguardar, estou gerando o arquivo .xlsx")
# lb_aviso.place(x=10, y=-10, width=300, height=100)

# btn = Button(avs, text="OK", command=avs.after(3000, lambda: avs.destroy())).pack

# #btn.place(x=135, y=60, width=30, height=30)
# avs.mainloop()
# #time.sleep(3)

# avs2 = Tk()
# avs2.title("Aviso")
# avs2.geometry("300x100")

# lb_aviso2 = Label(avs2,text="Arquivo Salvo!")
# lb_aviso2.place(x=10, y=-10, width=300, height=100)

# btn2 = Button(avs2, text="OK", command=avs2.destroy)
# btn2.pack()
# btn2.place(x=135, y=60, width=30, height=30)

# avs2.mainloop()

# notificacao = Notification(app_id="Aviso", title="Notificação", msg="Arquivo .xlsx Pronto!!!", duration="long")

# notificacao.show()
