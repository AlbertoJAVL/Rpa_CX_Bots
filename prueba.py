import os
import re

output = os.popen('wmic process get description, processid').read()
separador = output.split("\n")
print("limpiando")
for val in separador:
    # print(val)
    a = re.search("py.exe",val)
    b = re.search("python.exe",val)
    if a:
        os.system('taskkill /IM py.exe')
    if b:
        os.system('taskkill /IM python.exe')

