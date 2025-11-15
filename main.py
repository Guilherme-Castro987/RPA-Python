import webbrowser
import pyautogui
from time import sleep
import os
import pyperclip
from math import floor


controlador = 0
while controlador == 0:
    # Definindo o Link para acesso ao WMS
    url_wms = "https://wms-americanas.azr.internal.americanas.io/pedidos-abertos-fisico/pedidos-por-etapa-fisico?planta=UBER"

    # Iniciando o navegado e acessando o WMS Flash
    webbrowser.open(url_wms)
    sleep(25)
    # Iniciando os comandos no teclado para realizar o login
    pyautogui.press("tab")
    sleep(1)
    pyautogui.press("enter")
    sleep(10)
    pyautogui.write("coloca o login do associado",interval= 0.3)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("enter")
    sleep(5)
    pyautogui.write("coloca a senha do associado",interval= 0.3)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("enter")
    sleep(5)
    pyautogui.press("enter")
    sleep(25)
    pyautogui.press("tab")  
    sleep(1)
    pyautogui.press("enter")
    sleep(1)
    pyautogui.press("tab")
    sleep(1)
    pyautogui.press("enter")
    sleep(1)
    for x in range(7):
        pyautogui.press("tab")
    pyautogui.press("enter")
    sleep(5)
    try:
        botao_uberlandia = pyautogui.locateCenterOnScreen("cd.PNG")
        sleep(1)
        pyautogui.moveTo(botao_uberlandia,duration= 0.5)
        pyautogui.click()
        controlador = 1
    except:
        pyautogui.hotkey("alt","f4")
        sleep(3)
        controlador = 0

sleep(2)
pyautogui.scroll(-1000)
sleep(2)
#inicio #Point(x=29, y=347)
pyautogui.moveTo(x=29, y=347)
sleep(1)
pyautogui.mouseDown()
sleep(1)
# Fim Point(x=1629, y=896)
pyautogui.moveTo(x=1629, y=896)
sleep(1)
pyautogui.mouseUp()
sleep(1)
pyautogui.hotkey("ctrl","c")
sleep(2)
os.startfile("RupturaExpedição.xlsm")
sleep(50)
try:
    botao_habilitar = pyautogui.locateCenterOnScreen("habilitar.PNG")
    pyautogui.moveTo(botao_habilitar,duration= 0.5)
    pyautogui.click()
except:
    pass
pyautogui.press('enter')
sleep(40)
pyautogui.press("up")
pyautogui.press("up")
pyautogui.press("down")
sleep(1)
pyautogui.hotkey("ctrl","v")
sleep(1)
pyautogui.hotkey("ctrl","down")
sleep(1)
pyautogui.press("down")
sleep(1)
pyautogui.hotkey("alt","tab")
sleep(2)
pyautogui.scroll(-1000)
sleep(1)
botao_voltar = pyautogui.locateCenterOnScreen("anterior2.png",confidence=0.9)
pyautogui.moveTo(botao_voltar,duration=0.5)
sleep(1)
pyautogui.click()
pyautogui.click()
sleep(1)
pyautogui.hotkey("ctrl","c")
quantidade_lojas = int(pyperclip.paste())
divisao = quantidade_lojas / 10
if divisao.is_integer():
    paginas = int(divisao) - 1
else:
    paginas = floor(divisao)
sleep(1)
for y in range(paginas):
    pyautogui.scroll(-1000)
    sleep(1)
    botao_proximo2 = pyautogui.locateCenterOnScreen("proximo2.PNG",confidence= 0.9)
    pyautogui.click(botao_proximo2)
    sleep(2)
    pyautogui.moveTo(x=29, y=347)
    sleep(2)
    pyautogui.mouseDown()
    sleep(2)
    pyautogui.moveTo(x=1629, y=896)
    sleep(2)
    pyautogui.mouseUp()
    pyautogui.hotkey("ctrl","c")
    sleep(2)
    pyautogui.hotkey("alt","tab")
    sleep(2)
    pyautogui.press("up")
    sleep(1)
    pyautogui.press("down")
    sleep(1)
    pyautogui.hotkey("ctrl","v")
    sleep(1)
    pyautogui.hotkey("ctrl","down")
    sleep(1)
    pyautogui.press("down")
    sleep(1)
    pyautogui.hotkey("alt","tab")
    sleep(2)
pyautogui.hotkey("alt","tab")
sleep(2)
pyautogui.hotkey("ctrl","up")
sleep(2)
pyautogui.hotkey("ctrl","up")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("alt","down")
sleep(2)
for x in range(8):
    pyautogui.press("down")
    sleep(0.5)
pyautogui.write("loja",interval= 0.5)
sleep(2)
pyautogui.press("enter")
sleep(2)
pyautogui.press("left")
sleep(2)
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","shift","down")
sleep(2)
pyautogui.hotkey("ctrl","c")
sleep(2)
pyautogui.hotkey("ctrl","pagedown")
sleep(2)
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","v")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("ctrl","pageup")
sleep(2)
pyautogui.press("up")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("alt","down")
sleep(2)
for x in range(8):
    pyautogui.press("down")
    sleep(0.5)
pyautogui.write("Total Und's",interval= 0.5)
sleep(2)
pyautogui.press("enter")
sleep(2)
pyautogui.press("left")
sleep(2)
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","shift","down")
sleep(2)
pyautogui.hotkey("ctrl","c")
sleep(2)
pyautogui.hotkey("ctrl","pagedown")
sleep(2)
pyautogui.hotkey("ctrl","v")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("ctrl","pageup")
sleep(2)
pyautogui.press("up")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("alt","down")
sleep(2)
for x in range(8):
    pyautogui.press("down")
    sleep(0.5)
pyautogui.write("Entrada",interval= 0.5)
sleep(2)
pyautogui.press("enter")
sleep(2)
pyautogui.press("left")
sleep(2)
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","shift","down")
sleep(2)
pyautogui.hotkey("ctrl","c")
sleep(2)
pyautogui.hotkey("ctrl","pagedown")
sleep(2)
pyautogui.hotkey("ctrl","v")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("ctrl","pageup")
sleep(2)
pyautogui.press("up")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("alt","down")
sleep(2)
for x in range(8):
    pyautogui.press("down")
    sleep(0.5)
pyautogui.write("Picking",interval= 0.5)
sleep(2)
pyautogui.press("enter")
sleep(2)
pyautogui.press("left")
sleep(2)
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","shift","down")
sleep(2)
pyautogui.hotkey("ctrl","c")
sleep(2)
pyautogui.hotkey("ctrl","pagedown")
sleep(2)
pyautogui.hotkey("ctrl","v")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("ctrl","pageup")
sleep(2)
pyautogui.press("up")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("alt","down")
sleep(2)
for x in range(8):
    pyautogui.press("down")
    sleep(0.5)
pyautogui.write("Aguardando Conf.",interval= 0.5)
sleep(2)
pyautogui.press("enter")
sleep(2)
pyautogui.press("left")
sleep(2)
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","shift","down")
sleep(2)
pyautogui.hotkey("ctrl","c")
sleep(2)
pyautogui.hotkey("ctrl","pagedown")
sleep(2)
pyautogui.hotkey("ctrl","v")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("ctrl","pageup")
sleep(2)
pyautogui.press("up")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("alt","down")
sleep(2)
for x in range(8):
    pyautogui.press("down")
    sleep(0.5)
pyautogui.write("DANF solicitado",interval= 0.5)
sleep(2)
pyautogui.press("enter")
sleep(2)
pyautogui.press("left")
sleep(2)
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","shift","down")
sleep(2)
pyautogui.hotkey("ctrl","c")
sleep(2)
pyautogui.hotkey("ctrl","pagedown")
sleep(2)
pyautogui.hotkey("ctrl","v")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("ctrl","pageup")
sleep(2)
pyautogui.press("up")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("alt","down")
sleep(2)
for x in range(8):
    pyautogui.press("down")
    sleep(0.5)
pyautogui.write("DANF aprov",interval= 0.5)
sleep(2)
pyautogui.press("enter")
sleep(2)
pyautogui.press("left")
sleep(2)
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","shift","down")
sleep(2)
pyautogui.hotkey("ctrl","c")
sleep(2)
pyautogui.hotkey("ctrl","pagedown")
sleep(2)
pyautogui.hotkey("ctrl","v")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("ctrl","pageup")
sleep(2)
pyautogui.press("up")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("alt","down")
sleep(2)
for x in range(8):
    pyautogui.press("down")
    sleep(0.5)
pyautogui.write("SEC",interval= 0.5)
sleep(2)
pyautogui.press("enter")
sleep(2)
pyautogui.press("left")
sleep(2)
pyautogui.press("down")
sleep(2)
pyautogui.hotkey("ctrl","shift","down")
sleep(2)
pyautogui.hotkey("ctrl","c")
sleep(2)
pyautogui.hotkey("ctrl","pagedown")
sleep(2)
pyautogui.hotkey("ctrl","v")
sleep(2)
pyautogui.press("up")
sleep(2)
pyautogui.hotkey("ctrl","left")
sleep(2)
pyautogui.hotkey("ctrl","pageup")
sleep(2)
pyautogui.press("up")
sleep(2)
pyautogui.press("right")
sleep(2)
pyautogui.hotkey("alt","down")
sleep(2)
for x in range(5):
    pyautogui.press("down")
    sleep(0.5)
pyautogui.press("enter")
sleep(2)
pyautogui.press("left")
sleep(2)
pyautogui.hotkey("ctrl","b")
sleep(2)
pyautogui.hotkey("alt","f4")
sleep(15)
pyautogui.hotkey("alt","f4")

# Inicio para enviar os e-mail
    
import pandas as pd
import win32com.client as win32
from datetime import datetime

# Caminho do arquivo
arquivo = "RupturaExpedição.xlsm"

# ==============================
# 🔹 Leitura da aba "Viagens"
# ==============================
df_viagens = pd.read_excel(arquivo, sheet_name='Viagens')

# Filtra linhas com valor na coluna 'Placa'
df_viagens = df_viagens[df_viagens['Placa'].notna()]

# Padroniza 'Placa'
df_viagens['Placa'] = df_viagens['Placa'].astype(str).str.upper().str.strip()

# Converte colunas numéricas
if 'Viagem' in df_viagens.columns and pd.api.types.is_float_dtype(df_viagens['Viagem']):
    df_viagens['Viagem'] = df_viagens['Viagem'].astype("Int64")

if 'Turno' in df_viagens.columns and pd.api.types.is_float_dtype(df_viagens['Turno']):
    df_viagens['Turno'] = df_viagens['Turno'].fillna(0).astype(int)

colunas_totais = ['DropPoint', 'Eclusa', 'Aguardando SEC', 'SEC', 'Rup.Pend']
for col in colunas_totais:
    if col in df_viagens.columns and pd.api.types.is_float_dtype(df_viagens[col]):
        df_viagens[col] = df_viagens[col].fillna(0).astype(int)

# Reordena colunas
ordem_colunas = ['Data', 'Viagem', 'Placa', 'Status', 'Cluster', 'Turno',
                 'DropPoint', 'Eclusa', 'Aguardando SEC', 'SEC', 'Rup.Pend']
df_viagens = df_viagens[[col for col in ordem_colunas if col in df_viagens.columns]]

# Linha de total
linha_total = []
for col in df_viagens.columns:
    if col in colunas_totais:
        total = df_viagens[col].sum()
        linha_total.append(total)
    elif col == 'Cluster':
        linha_total.append('TOTAL')
    else:
        linha_total.append('')
df_total = pd.DataFrame([linha_total], columns=df_viagens.columns)
df_viagens = pd.concat([df_viagens, df_total], ignore_index=True)

# ==============================
# 🔹 Leitura da aba "Consolidado"
# ==============================
df_consolidado = pd.read_excel(arquivo, sheet_name='Consolidado', usecols="A:O")

# Remove linhas completamente vazias
df_consolidado = df_consolidado.dropna(how='all')

# Remove valores NaN (substitui por vazio para não aparecer no e-mail)
df_consolidado = df_consolidado.fillna("")

# 🔹 Converte colunas numéricas específicas para inteiro
colunas_int = ["Total und's", "Entrada", "Viagem", "Turno", "Ruptura"]

for col in colunas_int:
    if col in df_consolidado.columns:
        # Tenta converter; se falhar, ignora erro
        df_consolidado[col] = (
            pd.to_numeric(df_consolidado[col], errors='coerce')  # força números
            .fillna(0)  # substitui NaN por 0
            .astype(int)  # converte pra inteiro
        )

# ==============================
# 🔹 Estilo HTML compartilhado
# ==============================
estilo = """
<style>
    table {
        border-collapse: collapse;
        width: 100%;
        text-align: center;
        font-family: Arial, sans-serif;
        table-layout: auto;
    }
    thead th {
        background-color: #d3d3d3;
        border: 1px solid black;
        padding: 6px;
        white-space: nowrap;
    }
    tbody td {
        border: 1px solid black;
        padding: 6px;
        white-space: nowrap;
    }
    tbody tr:nth-child(even) {
        background-color: #f9f9f9;
    }
    tbody tr:nth-child(odd) {
        background-color: #ffffff;
    }
    tbody tr:last-child {
        background-color: #d3d3d3;
        font-weight: bold;
    }
</style>
"""

# ==============================
# 🔹 Monta as duas tabelas HTML
# ==============================
html_viagens = df_viagens.to_html(index=False, border=0, justify='center')
html_consolidado = df_consolidado.to_html(index=False, border=0, justify='center')

# ==============================
# 🔹 Corpo do e-mail
# ==============================
corpo_email = estilo + f"""
<p>Segue acompanhamento da expedição:</p>

<h3>🚚 Tabela de Viagens</h3>
{html_viagens}

<br><br>

<h3>📊 Tabela Consolidado</h3>
{html_consolidado}
"""

# ==============================
# 🔹 Envia via Outlook
# ==============================
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "larrire.marques@americanas.io;lucas.dimoura@americanas.io;lynconl.souza@americanas.io;ricafreitas@americanas.io;mailla.amorim@americanas.io;raphael.brandao@americanas.io;expedicao.cdmg@americanas.io;guilherme.dcastro@americanas.io;hajykaar.souza@americanas.io;luiz.asilva@americanas.io;braulio.terra@americanas.io"
mail.Subject = f"Acompanhamento Expedição - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
mail.HTMLBody = corpo_email
mail.Send()

print(f"✅ E-mail do relatório de viagens enviado com sucesso!{datetime.now().strftime('%d/%m/%Y %H:%M')}")

