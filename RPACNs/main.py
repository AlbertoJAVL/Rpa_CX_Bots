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

    try:
        # Eliminación de Temporales
        delTemporales()

        # Inicio de Sesion
        credenciales = api.get_orden_servicio2()
        driver, status_logueo = loginSiebel(credenciales['procesoUser'], credenciales['procesoPassword'])
        # driver, status_logueo = loginSiebel('bavila', 'Garciaydiego2025@')
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

                if 'PROMESA DE PAGO' in info['solucion'].upper(): solucion = 'PAGO COMPLETO'
                elif 'PROMESA CON PROMOCION'  in info['solucion'].upper(): solucion = 'PAGO CON PROMOCION'
                else: 
                    status = 'No aplica: Solucion Invalida'
                    response = api.ajusteCerrado(
                        info['id'],
                        '-',
                        info['fechaCaptura'],
                        info['fechaCompletado'],
                        status,
                        info['cve_usuario'],
                        ip,
                        info['cuenta'],
                        info['fechaSubida'],
                        info['categoria'],
                        info['mootivo'],
                        info['subMotivo'],
                        info['solucion'],
                        info['saldoIncobrable'],
                        info['promocion'],
                        info['ajuste'],
                        info['fechaGestion'],
                        info['tipo'])
                    print(response)
                    if resultado == False: return False

                if solucion == 'PAGO COMPLETO': comentario = f"SALDO INCOBRABLE: {info['saldoIncobrable']}\nFECHA: {info['fechaGestion']}\n{info['tipo']}\nCLIENTE CUENTA CON PROMOCION: ({info['promocion']})"
                else: comentario = f"SALDO INCOBRABLE: {info['saldoIncobrable']}\nPROMOCION: {info['promocion']}\nAJUSTE: {info['ajuste']}\nFECHA: {info['fechaGestion']}\n{info['tipo']}"

                comentario = comentario.replace('/', ' ').replace('%', ' ').replace('.', ' ').replace('$', ' ').replace('_', ' ').replace('-', ' ').replace(',', ' ')



                resultado, valEstado, numeroCN = inicio(driver, info['cuenta'], solucion, comentario)
                print(f'→ Resultado: {str(resultado)}')
                print(f'→ Estado Generado: {valEstado}')
                print(f'→ Numero CN Generado: {numeroCN}')

                status = valEstado
                response = api.ajusteCerrado(
                        info['id'],
                        numeroCN,
                        info['fechaCaptura'],
                        info['fechaCompletado'],
                        status,
                        info['cve_usuario'],
                        ip,
                        info['cuenta'],
                        info['fechaSubida'],
                        info['categoria'],
                        info['mootivo'],
                        info['subMotivo'],
                        info['solucion'],
                        info['saldoIncobrable'],
                        info['promocion'],
                        info['ajuste'],
                        info['fechaGestion'],
                        info['tipo'])
                print(response)
                if resultado == False: return False
                    

            else:
                try:
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
                info0['fechaSubida'],
                info0['categoria'],
                info0['mootivo'],
                info0['subMotivo'],
                info0['solucion'],
                info0['saldoIncobrable'],
                info0['promocion'],
                info0['ajuste'],
                info0['fechaGestion'],
                info0['tipo'])
            sleep(5)
            print('Eliminando Temporales')
            main()
        else: 
            sleep(5)
            print('------------------------ SIN REGISTROS ------------------------')
            os.system('cls')

    except: print('#########################\nERROR MAIN\n#########################')