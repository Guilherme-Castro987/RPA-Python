import webbrowser
from time import sleep
import pyautogui
import os
import pyperclip
from math import floor
import pyscreeze
import pygetwindow as gw
import pandas as pd
import win32com.client as win32
from datetime import datetime

# Abre o navegador para logar no WMS
controlador = 0
while controlador == 0:
    webbrowser.open("https://wms-americanas.azr.internal.americanas.io/pedidos-abertos-fisico/pedidos-por-etapa-fisico?planta=UBER")
    sleep(20)
    janela = gw.getActiveWindow()
    janela.moveTo(0,0)
    janela.maximize()
    sleep(6)
    # Entra no WMS
    pyautogui.press("tab")
    pyautogui.press("enter")
    sleep(20)
    # Entra no login
    pyautogui.write("guilherme.dcastro@americanas.io",interval= 0.1)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("enter")
    sleep(20)
    # Entra na senha
    pyautogui.write("@MiguelFrancisco041016")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("enter")
    sleep(10)
    # Aceita logar
    pyautogui.press("enter")

    sleep(55)
    # Na tela do WMS centraliza o mouse no meio da tela
    largura,altura = pyautogui.size()
    pyautogui.moveTo(largura/2,altura/2,duration= 0.5)
    sleep(55)
    pyautogui.scroll(-1000)
    sleep(3)
    try:
        Danf_aprovado = pyautogui.locateCenterOnScreen("danfAprovado.PNG",confidence=0.9)
        pyautogui.moveTo(Danf_aprovado,duration=0.5)
        controlador  = 1
    except:
        pyautogui.hotkey("alt","f4")
        controlador  = 0
Danf_aprovado = pyautogui.locateCenterOnScreen("danfAprovado.PNG",confidence=0.9)
pyautogui.moveTo(Danf_aprovado,duration=0.5)
sleep(2)
pyautogui.click()
sleep(2)
pyautogui.scroll(-1000)
sleep(3)
inicio = (307,333)
fim = (1692,874)
pyautogui.moveTo(inicio[0],inicio[1],duration=0.5)
pyautogui.mouseDown()
pyautogui.moveTo(fim[0],fim[1],duration=0.5)
pyautogui.mouseUp()
sleep(3)
pyautogui.hotkey("ctrl","c")
planilha_idades = r"idades.xlsm"
os.startfile(planilha_idades)
sleep(60)
controlador = 0
while controlador == 0:
    try:
        botao_limpar = pyautogui.locateCenterOnScreen("limpar.PNG",confidence=0.9)
        pyautogui.moveTo(botao_limpar,duration=0.5)
        controlador = 1
    except:
        pyautogui.hotkey("alt","f4")
        sleep(30)
        os.startfile(planilha_idades)
        sleep(60)
        controlador = 0
botao_limpar = pyautogui.locateCenterOnScreen("limpar.PNG",confidence=0.9)
pyautogui.moveTo(botao_limpar,duration=0.5)
sleep(2)
pyautogui.click()
sleep(0.5)
pyautogui.click()
sleep(0.5)
pyautogui.click()
sleep(8)
pyautogui.hotkey("alt","tab")
sleep(8)
pyautogui.scroll(-1000)
botao_anterior = pyautogui.locateCenterOnScreen("anterior.PNG",confidence=0.9)
pyautogui.moveTo(botao_anterior,duration=0.5)
sleep(1)
pyautogui.click()
pyautogui.click()
sleep(2)
pyautogui.hotkey("ctrl","c")
sleep(3)
quantidade_lojas = int(pyperclip.paste())
divisao = quantidade_lojas / 10
if divisao.is_integer():
    paginas = int(divisao) - 1
else:
    paginas = floor(divisao)

for x in range(paginas):
    try:
        botao_proximo = pyautogui.locateCenterOnScreen("proximo.PNG",confidence=0.9)
        pyautogui.moveTo(botao_proximo,duration=0.5)
        sleep(2)
        pyautogui.click()
        sleep(2)
        pyautogui.moveTo(inicio[0],inicio[1],duration=0.5)
        sleep(2)
        pyautogui.scroll(-1000)
        pyautogui.click()
        pyautogui.mouseDown()
        pyautogui.moveTo(fim[0],fim[1],duration=0.5)
        pyautogui.mouseUp
        sleep(1)
        pyautogui.hotkey("ctrl","c")
        sleep(1)
        pyautogui.hotkey("alt","tab")
        sleep(1)
        pyautogui.moveTo(x=1336,y=30)
        sleep(1)
        pyautogui.click()
        sleep(2)
        pyautogui.click()
        sleep(1)
        pyautogui.hotkey("ctrl","down")
        sleep(1)
        pyautogui.press("down")
        sleep(0.5)
        pyautogui.hotkey("ctrl","v")
        sleep(1)
        pyautogui.hotkey("alt","tab")
        sleep(5)
        pyautogui.moveTo(largura/2,altura/2,duration= 0.5)
        sleep(2)
        pyautogui.scroll(-1000)
        sleep(5)
    except:
        pyautogui.hotkey("alt","f4")
        sleep(10)
        pyautogui.hotkey("alt","f4")
        sleep(10)
        continue
pyautogui.hotkey("alt","tab")
sleep(1)
pyautogui.hotkey("ctrl","up")
sleep(2)
pyautogui.hotkey("ctrl","up")
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("alt","down")
for i in range(8):
    pyautogui.press("down")
    sleep(1)
sleep(2)
pyautogui.press("1")
sleep(2)
pyautogui.press("Enter")
sleep(2)
pyautogui.press("left")
sleep(2)
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","shift","down")
sleep(2)
pyautogui.hotkey("ctrl","c")
sleep(2)
pyautogui.press("up")
sleep(2)
pyautogui.hotkey("ctrl","pagedown")
sleep(2)
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","v")
sleep(8)
pyautogui.press("up")
sleep(2)
pyautogui.hotkey("ctrl","pageup")
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("alt","down")
for i in range(8):
    pyautogui.press("down")
    sleep(1)
sleep(2)
pyautogui.press("3")
sleep(2)
pyautogui.press("Enter")
sleep(2)
pyautogui.press("left")
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","shift","down")
sleep(2)
pyautogui.hotkey("ctrl","c")
sleep(2)
pyautogui.press("up")
sleep(2)
pyautogui.hotkey("ctrl","pagedown")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","v")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("ctrl","pageup")
sleep(2)
pyautogui.press("right")
pyautogui.hotkey("alt","down")
for i in range(8):
    pyautogui.press("down")
    sleep(1)
sleep(2)
pyautogui.press("4")
sleep(2)
pyautogui.press("Enter")
sleep(2)
pyautogui.press("left")
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","shift","down")
sleep(2)
pyautogui.hotkey("ctrl","c")
sleep(2)
pyautogui.press("up")
sleep(2)
pyautogui.hotkey("ctrl","pagedown")
sleep(2)
pyautogui.hotkey("ctrl","v")
sleep(8)
pyautogui.press("up")
sleep(2)
pyautogui.press("left")
pyautogui.press("left")
sleep(2)
pyautogui.hotkey("ctrl","pageup")
sleep(2)
pyautogui.press("right")
pyautogui.hotkey("alt","down")
for i in range(5):
    pyautogui.press("down")
    sleep(1)
sleep(2)
pyautogui.press("Enter")
sleep(2)
pyautogui.press("left")
pyautogui.hotkey("ctrl","b")
sleep(30)
pyautogui.hotkey("alt","f4")
sleep(50)
pyautogui.hotkey("alt","f4")
sleep(5)

data_hora_atual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Lê o arquivo, seleciona a aba e seleciona as colunas que vamos utilizar
df = pd.read_excel("idades.xlsm",sheet_name="Relação",usecols="A:D")

# Renomeia as colunas
df.columns = ["Loja","Und","Status","Cluster"]

# Lista das idades
status_list = ["Atrasado","D2","D1","D0"]

# Cria um dicionario para armazear cada dataframe filtrado
base_filtrada = {}

# Filtra e salva separadamente
for status in status_list:
    filtro = df[df["Status"]==status].sort_values(by="Cluster").copy()
    filtro["Und"] = filtro["Und"].astype("Int64")
    base_filtrada[status] = filtro

def df_to_html_inline(df):
    return df.to_html(
        index=False,
        border=1,
        escape=False,
        justify='center',
        classes='dataframe',
        render_links=True
    ).replace('<table border="1" class="dataframe">', '<table style="border-collapse: collapse; width: 100%;">') \
    .replace('<th>', '<th style="border: 1px solid black; padding: 5px; background-color: #f2f2f2; text-align: center;">') \
    .replace('<td>', '<td style="border: 1px solid black; padding: 5px; text-align: center;">')

# Montar o escopo do e-mail
html_corpo = "<p>Segue idades Expedição:</p>"
for status, base in base_filtrada.items():
    html_corpo += f"<h3>{status}</h3>"
    html_corpo += df_to_html_inline(base)

# Envia pelo Outlook
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.Subject = f"Relatório de Idades Expedição - {data_hora_atual}"
mail.To = "larrire.marques@americanas.io;lucas.dimoura@americanas.io;lynconl.souza@americanas.io;ricafreitas@americanas.io;mailla.amorim@americanas.io;raphael.brandao@americanas.io;expedicao.cdmg@americanas.io;guilherme.dcastro@americanas.io;hajykaar.souza@americanas.io;luiz.asilva@americanas.io;braulio.terra@americanas.io"  
mail.HTMLBody = html_corpo  # Usa o corpo com tabelas HTML
mail.Send()  # ou use mail.Display() para ver antes de enviar

# Loop de uma hora para não bloquear o PC

print(f"✅ E-mail das idades enviado com sucesso!{datetime.now().strftime('%d/%m/%Y %H:%M')}")
sleep(5)
#pyautogui.hotkey("alt","f4")