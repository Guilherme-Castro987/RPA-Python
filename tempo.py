import datetime
import time

while True:
    agora = datetime.datetime.now()
    
    # Se for 23h ou mais, encerra o loop
    if agora.hour >= 21:
        print("Encerrando execução. Já são 23:00 ou mais.")
        break

    # Coloque aqui o código principal que precisa rodar
    print(f"Executando... agora são {agora.strftime('%H:%M:%S')}")

    # Espera alguns segundos antes de repetir (evita uso excessivo de CPU)
    time.sleep(10)
   
import subprocess
import time

while True:
    print("Iniciando Idades_Expedicao.py...")
    subprocess.run(["python", "Idades_Expedicao.py"])

    print("Iniciando Ruptura_Expedicao.py...")
    subprocess.run(["python", "Ruptura_Expedicao.py"])

    # (Opcional) Aguarda alguns segundos antes de repetir
    time.sleep(5)