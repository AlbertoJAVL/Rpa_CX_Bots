from json.decoder import JSONDecodeError
import requests
import json
from time import sleep
import socket

hostname = socket.gethostname()
ip = socket.gethostbyname(hostname)
print(ip)


url = 'https://rpabackizzi.azurewebsites.net/AjustesCambiosServicios/getCuentaRetencion'
urlUpdate = 'https://rpabackizzi.azurewebsites.net/AjustesCambiosServicios/ActualizaRetencion'
urlPassUser = f'https://rpabackizzi.azurewebsites.net/Bots/getProcessRetencion?ip={str(ip)}'

def get_orden_servicio():
    try:
        response = requests.get(url)
        if response.status_code == 200:
            responseApi = json.loads(response.text)
            print('API CORRECTA')
            print(responseApi)
            return responseApi
        elif response.status_code == 401: return print("Anauthorized")
        elif response.status_code == 404: return print("Not Found")
        elif response.status_code == 500: return print("Internal Server Error")
        
    except JSONDecodeError: return response.body_not_json

def get_orden_servicio2():
    try:
        response = requests.get(urlPassUser)
        if response.status_code == 200:
            responseApi = json.loads(response.text)
            print('API CORRECTA')
            print(responseApi)
            return responseApi
        elif response.status_code == 401: return print("Anauthorized")
        elif response.status_code == 404: return print("Not Found")
        elif response.status_code == 500: return print("Internal Server Error")
        
    except JSONDecodeError: return response.body_not_json

def update(datos, parametros):

    try:
        response = requests.put(urlUpdate, params=parametros, json=datos)
        if response.status_code == 200:
            responseApi = json.loads(response.text)
            print('ACTUALIZADO')
            return responseApi
        
        elif response.status_code == 401: return print("Anauthorized")
        elif response.status_code == 404: return print("Not Found")
        elif response.status_code == 500: return print("Internal Server Error")

    except JSONDecodeError: return response.body_not_json


def ajusteCerrado(id, numeroAjuste, fechaCaptura, fechaCompletado, status, cve_usuario, ip, cuenta, casoNegocio, proceso, montoAjustado, ajustePrevio, sinEquipo):
    datos = {
        'id' : id,
        'numeroAjuste' : numeroAjuste,
        'fechaCaptura' : fechaCaptura,
        'fechaCompletado' : fechaCompletado,
        'status' : status,
        'cve_usuario' : cve_usuario,
        'ip' : ip,
        'cuenta' : cuenta,
        'casoNegocio' : casoNegocio,
        'proceso' : proceso,
        'montoAjustado' : montoAjustado,
        'ajustePrevio' : ajustePrevio,
        'equipo' : sinEquipo
    }

    parametros = { 'id' : id }
    return update(datos, parametros)

