from os import environ, path, remove, listdir
import apiCyberHubOrdenes as api
from login import loginSiebel
from shutil import rmtree
from funcionalidad import *
import socket
import os

def delTemporales():

    temp_folder = environ['TEMP']

    try:

        temp_files = listdir(temp_folder)

        for temp_file in temp_files:
            temp_file_path = path.join(temp_folder, temp_file)

            try:
                if path.isfile(temp_file_path): remove(temp_file_path)
                elif path.isdir(temp_file_path): rmtree(temp_file_path)

            except: pass

        print('Eliminacion temporales OK!')

    except Exception as e: print('Se produjo un error al eliminar los temporales')

def main():
    print('main iniciado')
    
    try:
        # Eliminación de Temporales
        delTemporales()

        # Inicio de Sesion
        credenciales = api.get_orden_servicio2()
        driver, status_logueo = loginSiebel(credenciales['procesoUser'], credenciales['procesoPassword'])
        # driver, status_logueo = loginSiebel('ftorresma', 'Tobias1949++')
        if status_logueo == False: 
            print('→ LOGGIN INCORRECTO ←')
            return False

        while True:
            apiResponse = api.get_orden_servicio()
            info = apiResponse[0]
            print(info)
            host = socket.gethostname()
            ip = socket.gethostbyname(host)

            if info != 'SIN INFO':

                api.ajusteCerrado(
                    info['id'],
                    '-',
                    info['fechaCaptura'],
                    info['fechaCompletado'], 
                    'Procesando',
                    info['cve_usuario'],
                    ip,
                    info['cuenta'],
                    info['casoNegocio'], 
                    info['proceso'], 
                    '-', 
                    '-', 
                    info['equipo'])

                if info['equipo'] == None or info['equipo'] == '': sinEquipo = False
                elif info['equipo'].upper() == 'CC': sinEquipo = 'CC'
                elif info['equipo'].upper() == 'SI': sinEquipo = False
                else: sinEquipo = True

                resultado, numeroAjuste, montoAjuste, ajustePrevio = inicio(driver, credenciales['procesoUser'], info['cuenta'], info['casoNegocio'], info['proceso'], sinEquipo)
                print(f'→ Resultado: {str(resultado)}')
                print(f'→ Numero Ajuste: {str(numeroAjuste)}')
                print(f'→ monto Ajuste: {str(montoAjuste)}')
                
                
                if 'Inconsistencia' in resultado or 'No aplica' in resultado or 'pendiente' in resultado or 'Error' in resultado:
                    status = resultado
                    response = api.ajusteCerrado(info['id'],str(numeroAjuste),info['fechaCaptura'],info['fechaCompletado'],status,info['cve_usuario'],ip,info['cuenta'],info['casoNegocio'], info['proceso'], str(montoAjuste), str(ajustePrevio), info['equipo'])
                    print(response)
                    return False
                else:
                    resultado2 = cierreActividad(driver, info['cuenta'], info['casoNegocio'], info['proceso'])
                    print(f'→Resultado2: {str(resultado2)}')
                    if 'Inconsistencia' in resultado2 or 'No aplica' in resultado2 or 'Error' in resultado2:
                        status = resultado2
                        response = api.ajusteCerrado(info['id'],str(numeroAjuste),info['fechaCaptura'],info['fechaCompletado'],status,info['cve_usuario'],ip,info['cuenta'],info['casoNegocio'], info['proceso'], str(montoAjuste), str(ajustePrevio), info['equipo'])
                        print(response)
                        return False
                    else:
                        status = 'Cerrado'
                        print('→Generando Cambio de Status')
                        response = api.ajusteCerrado(info['id'],str(numeroAjuste),info['fechaCaptura'],info['fechaCompletado'],status,info['cve_usuario'],ip,info['cuenta'],info['casoNegocio'], info['proceso'], str(montoAjuste), str(ajustePrevio), info['equipo'])
                        print(response)
                    

            else:
                try:

                    driver.close()
                    return False
                    os.system('cls')
                    print('Esperando mas CN')
                    sleep(15)
                    print('Regreso a HOME')
                    home(driver)
                    print('##############\n FIN DE CICLO COMPLETO \n##############')

                except Exception: return False
    except: return False

while True == True:

    os.system(f'taskkill /f /im chrome.exe')
    os.system(f'taskkill /f /im chrome.exe')
    os.system(f'taskkill /f /im chrome.exe')

    try:
        apiResponse0 = api.get_orden_servicio()
        info0 = apiResponse0[0]
        print(info0)
        host = socket.gethostname()
        ip = socket.gethostbyname(host)

        if info0 != 'SIN INFO':

            api.ajusteCerrado(
            info0['id'],
            '-',
            info0['fechaCaptura'],
            info0['fechaCompletado'], 
            'Registro pendiente',
            info0['cve_usuario'],
            ip,
            info0['cuenta'],
            info0['casoNegocio'], 
            info0['proceso'], 
            '-', 
            '-', 
            info0['equipo'])
            sleep(5)
            print('Eliminando Temporales')
            main()
        else: 
            sleep(5)
            print('------------------------ SIN REGISTROS ------------------------')
            os.system('cls')

    except: print('#########################\nERROR MAIN\n#########################')


