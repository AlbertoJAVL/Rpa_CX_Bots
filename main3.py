from time import sleep
from dotenv import load_dotenv
import sys,os,re
import requests
# import #socket_boot as #socket
import socket



ID = ""
STATUS = ""


def changeStatusBot(estado):
    try:
        hostname= socket.gethostname()
        ip = socket.gethostbyname(hostname)
        url = f'https://rpabackizzi.azurewebsites.net/AjustesCambiosServicios/updateProcessStatusBotsRetencion?ip={ip}&estado={estado}'
        response = requests.get(url)

        # Verifica si la solicitud fue exitosa (código de estado 200)
        response.raise_for_status()

        # Accede a los datos de la respuesta (pueden ser JSON, texto, etc.)
        # print("Contenido de la respuesta:", response.text)
        return response.text


    except requests.exceptions.RequestException as e:
        print(f"Error en la solicitud: {e}")
        return None
        

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def stop():
    output = os.popen('wmic process get description,processid').read()
    separador = output.split('\n')
    for val in separador:
        # print(val)
        b = re.search("powershell.exe",val)
        # c = re.search("siebel.exe",val)
        c = re.search("chrome.exe",val)
        d = re.search("py.exe",val)
        if b:
            os.system('TASKKILL /F /IM powershell.exe /T')
            os.system('taskkill /IM chrome.exe')
            os.system('taskkill /IM py.exe')

while True:
    sleep(3)
    load_dotenv(override=True)
    id_1=os.getenv('proceso')
    status_1=os.getenv('status')
    print(id_1)
    print(status_1)
    if ID!= id_1 or STATUS!= status_1:
        ID = id_1
        STATUS = status_1

        if ID == "1" or ID == "33":
            
            print(f"Iniciando el proceso {STATUS} Retencion 0")
            stop()
            if STATUS=="STOPED": changeStatusBot(0)
            elif STATUS=="STARTED":
                sleep(5)
                aa ='./RPA3ESC/tele.py'
                os.system(f"start powershell python {aa}")
                changeStatusBot(1)

        elif ID == "5" or ID == "9":

            print(f"Iniciando el proceso {STATUS} Cargo Extemporaneo")
            stop()
            if STATUS=="STOPED": changeStatusBot(0)
            elif STATUS=="STARTED":
                sleep(5)
                aa ='./Rpa_cargoExt_convenio_cob/tele.py'
                os.system(f"start powershell python {aa}")
                changeStatusBot(1)

        elif ID == "10" or ID == "11":
            
            print(f"Iniciando el proceso {STATUS} Generacion CN's")
            stop()
            if STATUS=="STOPED": changeStatusBot(0)
            elif STATUS=="STARTED":
                sleep(5)
                aa ='./RPACNs/tele.py'
                os.system(f"start powershell python {aa}")
                changeStatusBot(1)

    else: print("No han cambiado las variables")