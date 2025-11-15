import subprocess
from time import sleep
import datetime
import pyautogui


# Horário para parar o loop
horario_inicio = datetime.time(9,10)
horario_fim = datetime.time(23,30)
# Loop para manter os 2 robôs funcionando
while True:
    agora = datetime.datetime.now().time()
    # Verificar se o horário limite chegou e
    if agora >= horario_inicio and agora <= horario_fim:
        subprocess.run(["python","wmsFlash.py"]) #robo das idades
        sleep(5)
        subprocess.run(["python","main.py"]) #robo das viagens 
        # Loop de uma hora para não bloquear o PC
        for x in range(50):
            pyautogui.press("down")
            sleep(1)
            pyautogui.press("up")
            sleep(30)
    else:
        for x in range(20):
            pyautogui.press("down")
            sleep(1)
            pyautogui.press("up")
            sleep(30)
