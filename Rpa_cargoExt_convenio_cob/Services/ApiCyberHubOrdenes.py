import requests
import json
import os 
from json.decoder import JSONDecodeError
from time import sleep
import socket

hostname= socket.gethostname()
ip = socket.gethostbyname(hostname)

print(ip)
#region URL

url = 'https://rpabackizzi.azurewebsites.net/AjustesCambiosServicios/getCuentaAjustesCambioServicio'
urlOrden = 'https://rpabackizzi.azurewebsites.net/AjustesCambiosServicios/ActualizaAjustesCambioServicio'
urlPassUser = f'https://rpabackizzi.azurewebsites.net/Bots/getProcessRetencion?ip={ip}'
#url_space ="https://api.ocr.space/parse/image"
#endregion




'''OBTIENE LA ORDEN DE SERVICIO
Recibe: 
        NADA
Entrega:
        Un código de respuesta o los datos de una orden

'''

def get_user():
    try:
        response = requests.get(urlPassUser) 
        if response.status_code == 200:
            reseponseApi = json.loads(response.text)
            print(reseponseApi)
            return reseponseApi

        elif response.status_code == 401:
            return print("Unauthorized")

        elif response.status_code == 404:
            return print("Not Found")

        elif response.status_code == 500:
            return print("Internal Server Error")

    except JSONDecodeError:
        return response.body_not_json

p = get_user()
print(p)

def get_orden_servicio():
    try:
        response = requests.get(url) 
        if response.status_code == 200:
            reseponseApi = json.loads(response.text)
            print('API correcta')
            return reseponseApi

        elif response.status_code == 401:
            return print("Unauthorized")

        elif response.status_code == 404:
            return print("Not Found")

        elif response.status_code == 500:
            return print("Internal Server Error")

    except JSONDecodeError:
        return response.body_not_json
# p = get_orden_servicio()
# print(p)

'''ACTUALIZACIÓN DE DATOS DE
Recibe:  
        Datos (previamente estructurados)
Devuelve:
        Un código de respuesta
'''

def update(datos, parametros):

    try:
        response = requests.put(urlOrden, params=parametros, json=datos, verify=False)
        if response.status_code == 200:
            responseApi = json.loads(response.text)
            print('Actualizado')
            return responseApi

        elif response.status_code == 401:
            return print("Unauthorized")

        elif response.status_code == 404:
            return print("Not Found")

        elif response.status_code == 500:
            return print("Internal Server Error")

    except JSONDecodeError:
            return response.body_not_json

'''CUANDO A UNA ORDEN SE LE HACE EL PROCESO
Recibe: 
        id_lead(INTEGER) & source_id (STRING)
Devuelve: 
        Llama a UPDATE para actualizar los datos en la base
'''

def ajusteCerrado(id, cuenta, casoNegocio, cve_usuario, motivo, motivoCliente, estado, status, procesando, ip, fechaCreado, fechaCompletado, fechaCarga, fechaVencimiento):
    datos = {
        'id' : id,
        'cuenta': cuenta,
        'casoNegocio' : casoNegocio,
        'cve_usuario' : cve_usuario,
        'motivo' : motivo,
        'motivoCliente' : motivoCliente,
        'estado' : estado,
        'status' : status,
        'procesando' : procesando,
        'ip' : ip,
        'fechaCreado' : fechaCreado,
        'fechaCompletado' : fechaCompletado,
        'fechaCarga' : fechaCarga,
        'fechaVencimiento' : fechaVencimiento
        }

    parametros = { 'id' : id }
    return update(datos, parametros)

# api = get_orden_servicio2()
# # info = api[0]
# print(info)
# ajusteCerrado(info['id'], info['cuenta'], info['casoNegocio'], info['cve_usuario'], info['motivo'], info['motivoCliente'], info['estado'], 'Cerrado', '1', 'vix10-PC', info['fechaCreado'], info['fechaCompletado'], info['fechaCarga'], info['fechaVencimiento'])

'''LA CUENTA NO COINCIDE CON LA DE SEIEBEL
Recibe: 
        id_lead(INTEGER) & source_id (STRING)
Devuelve: 
        Llama a UPDATE para actualizar los datos en la base
'''

def cuenta_erronea(id_lead,source_id):
    datos = {
        "source_id":source_id,
        "lead_id":id_lead,
        "error": False,
        "code": 3
        }                     
    return update(datos)


'''CUANDO A UNA ORDEN SE LE HACE EL PROCESO
Recibe: 
        id_lead(INTEGER) & source_id (STRING)
Devuelve: 
        Llama a UPDATE para actualizar los datos en la base
'''

def cuenta_trabaja_izzi(id_lead,source_id):
    datos = {
        "source_id":source_id,
        "lead_id":id_lead,
        "error": False,
        "code": 4
        }                      
    return update(datos)

'''CUANDO NO COINCIDENTE CON LA DE SIEBEL
Recibe: 
        id_lead(INTEGER) & source_id (STRING)
Devuelve: 
        Llama a UPDATE para actualizar los datos en la base
'''

def cuenta_errorsiebel(id_lead,source_id):
    datos = {
        "source_id":source_id,
        "lead_id":id_lead,
        "error": False,
        "code": 5
        }                      
    return update(datos)

'''API SIN DATOS
Recibe: 
        id_lead(INTEGER) & source_id (STRING)
Devuelve: 
        Llama a UPDATE para actualizar los datos en la base
'''

def api_sin_datos():
    datos = {
            "lead_id":0,
            "error": False,
            "code": 4}                       
    return update(datos)

#------------------------------------------------------------------
''' ACTUALIZA DATOS DEL CASO DE NEGOCIO
Recibe: 
        datos (previamente estructurados)
Devuelve: 
        Llama a UPDATE para actualizar los datos en la base
'''

def update_cn(datos):

    try:
        response = requests.put(urlCN, data=datos)
        if response.status_code == 200:
            responseApi = json.loads(response.text)
            return responseApi

        elif response.status_code == 401:
            return print("Unauthorized")

        elif response.status_code == 404:
            return print("Not Found")

        elif response.status_code == 500:
            return print("Internal Server Error")

    except JSONDecodeError:
            return response.body_not_json

''' CUANDO EL CASO DE NEGOCIO SE CERRÓ
Recibe: 
        datos (previamente estructurados)
Devuelve: 
        Llama a UPDATE_cn para actualizar los datos en la base
'''

def caso_negocio_cerrado(id_lead,source_id,cn):
    datos = {
        "source_id":source_id,
        "lead_id":id_lead,
        "cn":cn,
        "error": False,
        "code": 1
        }                      
    return update_cn(datos)

''' CUANDO EL CASO DE NEGOCIO YA EXISTE
Recibe: 
        datos (previamente estructurados)
Devuelve: 
        Llama a UPDATE_cn para actualizar los datos en la base
'''

def caso_negocio_existe(id_lead,source_id,cn):
    datos = {
        "source_id":source_id,
        "lead_id":id_lead,
        "cn":cn,
        "error": True,
        "code": 2
        }                      
    return update_cn(datos)


'''ENCUENTRA EL USUARIO SEGÚN LA ORDEN
Se puede decir que es un validador de que haya ordenes
'''
def usuario():
    # Si hay ordenes...
    orden = get_orden_servicio()
    #print(orden)
    status = orden['status']
    #print(status)
    if status == True:
        #Busca las credenciales
        keys = orden['access_data']
        return keys #Regresa user yu pass
    else:
        #En caso de que no haya ordenes...       
        #os.system("taskkill /im firefox.exe")
        print("No hay ordenes")
        print("Espere....")
        sleep(10)
        print('▬▬Reiniciando programa▬▬')
        usuario()
        
