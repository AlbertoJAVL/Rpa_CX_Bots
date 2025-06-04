from time import sleep
import requests
from datetime import datetime
import pytz
import autoit
import os
import re
import socket

count = 0
hostname= socket.gethostname()
ip = socket.gethostbyname(hostname)
count = 0
token = '6819354375:AAFb2UuBWfbOkT83YDyt2IH_lHSUgOpnkuU'
username = 'Neocall_Bot'


def send_msg():
    IST = pytz.timezone('America/Mexico_City')
    raw_TS = datetime.now(IST)
    curr_date= raw_TS.strftime("%d-%m-%Y")
    curr_time= raw_TS.strftime("%H-%M-%S")
    msg = f" {str(ip)} se ha detenido:  {curr_date} at {curr_time}"
    telegram_api=f"https://api.telegram.org/bot6819354375:AAFb2UuBWfbOkT83YDyt2IH_lHSUgOpnkuU/sendMessage?chat_id=-1002094293899&text={msg}"
    tel_resp = requests.get(telegram_api)
    if tel_resp.status_code == 200:
        print("INFO: Notification has been sent on Telegram")
    else:
        print("ERROR: Could not send Message")

def close_explorer():
    mas=0
    output = os.popen('wmic process get description, processid').read()
    separador = output.split("\n")
    for val in separador:
        a = re.search("iexplore.exe",val)
        if a:
            mas+=1
    if mas !=0:
        for i in range(mas-1):
            os.system('taskkill /IM Iexplore.exe')


def ini_bot():
    while True:
        count=0
        sleep(10)
        # a = autoit.process_exists("cmd.exe")
        a = autoit.process_exists("py.exe")
        # print(a)
        if a ==0:
            send_msg()
            os.startfile(r"C:\Rpa_CX_Bots\RPA3ESC\Ajustes2.bat")
           

try:
    ini_bot()
	
except KeyboardInterrupt:
    ini_bot()
    print("Keyboardinterrupt exception is caught")
    print(KeyboardInterrupt)


# output = os.popen('wmic process get description, processid').read()
# separador = output.split("\n")
# for val in separador:
#     print(val)
#     a = re.search("iexplore.exe",val)
#     if a:
#         count+=1
# print("Esto  es count ",count)

# autoit.win_close("iexplore.exe")
# a = autoit.process_exists("iexplore.exe")
# print(a)

# os.system('taskkill /F /IM Iexplore.exe')
# os.system('taskkill /IM Iexplore.exe')
# os.system('taskkill/im Iexplore.exe')
# b =os.system('tasklist')
# print(b)

# autoit.run("cmd.exe")
# autoit.send("cd c://PROYECTOS {ENTER}")
# autoit.send(" python tele.py")

# output = os.popen('wmic process get description, processid').read()
# print(output)

# separador = output.split("\n")
# for val in separador:
#     a = re.search("cmd.exe",val)
#     if a:
#         count+=1
# print("Esto es el count",count)

# def ini_bot():
#     cmd =0
#     while True:
#         cmd=0
#         sleep(5)
#         output = os.popen('wmic process get description, processid').read()
#         separador = output.split("\n")
#         for val in separador:
#             b = re.search("cmd.exe",val)
#             if b:
#                 cmd+=1
#         print("Esto es el cd",cmd)
        


# try:
#     ini_bot()

# except KeyboardInterrupt:
#     print(KeyboardInterrupt)



