#----------Selenium--------------------#
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
import autoit as it
import re
#-------------System-------------------#
from time import sleep
import os
import win32clipboard as cp
from datetime import *
#---------Mis funciones---------------#
from utileria import *
from logueo import *
import Services.ApiCyberHubOrdenes as api
from rutas import *
from IA import *

def ordenar(driver, engrane, opcionOrden, input, elementoOrden, orden, confirmar):

    ########################################################################

                ### FUNCION PARA ORDENAR TABLAS ###

    ########################################################################

    print('Ordenamiento de tabla')

    driver.find_element(By.XPATH, engrane).click() #Click sobre engrane
    sleep(1)
    driver.find_element(By.XPATH, opcionOrden).click() #Click sobre opcion "Ordenar"
    sleep(1)
    driver.find_element(By.XPATH, input).click() #Click sobre input "Ordenar Por"
    
    inputOrdenarPor = driver.find_element(By.XPATH, input) #Obtencion input "Ordenar Por"
    inputOrdenarPor.clear()
    inputOrdenarPor.send_keys(elementoOrden)
    inputOrdenarPor.send_keys(Keys.ESCAPE)

    sleep(1)

    driver.find_element(By.XPATH, orden).click() #Click sobre opcion "Descendiente"
    driver.find_element(By.XPATH, confirmar).click() #Click sobre boton "Aceptar"

    inputOrdenarPor = driver.find_element(By.XPATH, input) #Obtencion input "Ordenar Por"
    inputOrdenarPor.clear()
    inputOrdenarPor.send_keys(elementoOrden)
    inputOrdenarPor.send_keys(Keys.ESCAPE)

    sleep(1)

    driver.find_element(By.XPATH, orden).click() #Click sobre opcion "Descendiente"
    driver.find_element(By.XPATH, confirmar).click() #Click sobre boton "Aceptar"

def validacion_cuenta_fallas_servicio(driver, no_cuenta, no_caso, numero_pago):
    '''
    Funcion que hace el ajuste por Cargo Extemporaneo
    args:
        - driver
        - no_cuenta
        - no_caso

    out:
        - Bool: True en caso de que la cuenta cumpla con las validaciones para el ajuste
        - causa_rechazo (str): causa de validacion fallida

    '''
    try:
        wait = WebDriverWait(driver, 120)
        act = webdriver.ActionChains(driver)
        text_box('AJSUTE POR CONVENIO COBRANZA', '▬')
        status_pantalla_unica = pantalla_unica_consulta(driver, no_cuenta) 
        if status_pantalla_unica == False:
            text_box('Cuenta no valida', '▬')
            return False , 'Cuenta no valida', '',''
        
        print('▬ Inician validaciones de cuenta ▬')
        
        print('▬ Tipo Cuenta')
        tipo_cuenta = my_copy_by_xpath(driver, pantalla_unica['tipo_cuenta']['xpath'])
        if tipo_cuenta.upper() not in ['RESIDENCIAL', 'NEGOCIO']:
            error = 'tipo cuenta ' + tipo_cuenta
            print('Error ', error)
            return False, True, '', error
        

        print('▬ Subtipo Cuenta') 
        subtipo_cuenta = my_copy_by_xpath(driver, pantalla_unica['subtipo_cuenta']['xpath'])
        if subtipo_cuenta.upper() not in ['NORMAL', 'IZZI', 'EMPLEADOS IZZI']:
            error = 'subtipo cuenta ' + subtipo_cuenta
            print('Error ', error)
            return False, True, '', error


        ######### Validacion Caso de Negocio

        ######### Fin validacion

        print('▬ Ordenes de Servicio')

        ordenar(driver, engrane, opcionOrden, input, elementoOrden, orden, confirmar)
        
        tipoOS = driver.find_element(By.XPATH, ).click() #Click sobre "Tipo" en Ordenes de Servicio (Troublecall)
        tipoOS = my_copy(driver)
        if 'Trouble Call' in tipoOS or 'Trouble Call Video' in tipoOS:
            print('Orden de Servicio CORRECTA')
            fechaOrden = driver.find_element(By.XPATH, ).click() #Click sobre "Fecha de la Orden" fecha dentro del mismo mes de la solicitud
            fechaOrden = datetime.strptime(fechaOrden, '%d/%m/%Y %H:%M:%S')
            fechaOrden = fechaOrden.strftime('%Y-%m-%d %H:%M:%S')

            now = datetime.now()
            if (fechaOrden.month == now.month) and (fechaOrden.year == now.year):
                print('Orden dentro del mismo mes a la solicitud')
                estadoOS = driver.find_element(By.XPATH, ).click() #Click sobre "Estado" en Ordenes de Servicio (Troublecall)
                estadoOS = my_copy(driver)
                
                if 'Completa' in estadoOS:
                    print('Orden De Servicio Correcta')
                elif 'Cancelada' in estadoOS:
                    print('Orden de Servicio Cancelada')
                    return False, 'Cancelar'
                else:
                    error = 'Orden Servicio Incorrecta'
                    return False, error
            else:
                error = 'Fecha Orden fuera de mes'
                return False, error
        else:
            error = 'Orden Servcio Incorrecta'
            return False, error

        print('▬ Caso de Negocio')

        ordenar(driver, engrane, opcionOrden, input, elementoOrden, orden, confirmar)


        
    


        print('Validaciones correctas')
        return True, 'listo', ajuste, ''

    except Exception as e:
        print(f'ERROR ajustando el cargo extemporaneo. Caso NO. {no_caso}. CUENTA: {no_cuenta}')
        description_error('11','validacion_cuenta_cargo_extemporaneo',e)
        print('Error ', e)
        return False, str(e), '', 'Excepcion'
