#----------Selenium--------------------#
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.alert import Alert
import autoit as it
#-------------System-------------------#
from time import sleep
import os
import win32clipboard as cp
import  datetime
from pathlib import Path
#---------Mis funciones---------------#
from utileria import *
from logueo import *
import Services.ApiCyberHubOrdenes as api
from rutas import *


#################### Nuevas Librerias
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import os
import re
import shutil

#---------Variables globales---------------#

FORMATO_FECHA = "%Y/%m/%d %H:%M"
FORMATO_FECHA_INVERT = "%d/%m/%Y %H:%M"

def ordenar(driver, engrane, opcionOrden, input, elementoOrden, orden, confirmar):

    ########################################################################

                ### FUNCION PARA ORDENAR TABLAS ###

    ########################################################################
    try:
        print('Ordenamiento de tabla')

        driver.find_element(By.XPATH, engrane).click() #Click sobre engrane
        sleep(3)
        driver.find_element(By.XPATH, opcionOrden).click() #Click sobre opcion "Ordenar"
        sleep(6)
        driver.find_element(By.XPATH, input).click()
        sleep(2)
        inputOrdenarPor = driver.find_element(By.XPATH, input + '/input') #Obtencion input "Ordenar Por"
        sleep(2)
        inputOrdenarPor.clear()
        sleep(3)
        inputOrdenarPor.send_keys(elementoOrden)
        sleep(2)
        inputOrdenarPor.send_keys(Keys.ESCAPE)

        sleep(3)
        descendiente = driver.find_element(By.XPATH, orden) #Click sobre opcion "Descendiente"
        descendiente.click()
        descendiente.click()
        descendiente.send_keys(Keys.RETURN)
        # sleep(4)

        # try:
        #     driver.find_element(By.XPATH, confirmar).click() #Click sobre boton "Aceptar"
        # except Exception:
        #     driver.find_element(By.XPATH, confirmar).click() #Click sobre boton "Aceptar"

        print('Tabla ordenada')
    except Exception as e:
        print(e)
        print('------------- Fallo al ordenar la tabla ---------------')
        sleep(1000)

def actividades_asigandas(driver):
    '''
    Funcion que consulta las activdades asignadas y extrae una cuenta con un caso de negocio

    args:
        - driver (obj): instancia del navegador

    out:
        - TUPLA: (BOOL, 'INFORMACION DEL ERROR) True si se pudo ingresar a 'Actividades' y las ordeno
        - no_cuenta (str)
        - no_caso(str)
        - COMENTARIOS
    '''
    #Se define el maximo de espera para los elementos, si esto se excede, se genera una excepción
    try:
        wait = WebDriverWait(driver, 120)
        act = webdriver.ActionChains(driver)
        text_box('Pantalla::  Actividades (Activdiades asignadas)', '☼')
        
        #Espera a que la página tenga el titulo de Home 
        element = wait.until(EC.title_contains('inicial'))

        #Entra a la ventana de Actividades
        status_open = open_item_selenium_wait(driver, xpath =  home['actividades']['xpath'])
        if status_open == False:
            description_error('03','pantalla_Avtividades','No se pudo entrar a la pagina de Actividades')
            return status_open, []
        print('Pantalla correcta')

        #MIS ACTIVIDADES
        # open_item_selenium_wait(driver, xpath = actividades['menu_actividades']['xpath'] )
        # open_item_selenium_wait(driver, xpath = actividades['mis_actividades']['xpath'] )
        # status_open = open_item_selenium_wait(driver, xpath = actividades['busqueda_actividades']['xpath'] )
        # if status_open == False:
        #     description_error('03','mis_actividades','No se pudo seleccionar "Mis actividades"')
        #     return status_open
        # print('Mis actividades')
        
        ## SE FILTRAN POR FECHA Y ASIGNADAS

        open_item_selenium_wait(driver, xpath = '//*[@id="s_2_1_17_0_Ctrl"]' ) #LUPA
        it.send("{TAB 9}")
        sleep(2)
        it.send("ASIGNADA")
        sleep(2)
        it.send("{ENTER 2}")
        sleep(10)

        #COPIA LOS CAMPOS CUENTA, COMENTARIOS Y CASO
        it.send("+{TAB 2}")
        no_cuenta = my_copy(driver) #Valida que haya cuentas
        if no_cuenta == False or no_cuenta == '': return  (False, 'No hay cuentas'), '' , '' , '' 
        sleep(2)

        it.send("{TAB 4}")
        no_caso  = my_copy(driver)
        sleep(2)

        it.send("{TAB 1}")
        vencimiento  = my_copy(driver)
        vencimiento = datetime.datetime.strptime(vencimiento, FORMATO_FECHA_INVERT)

        fecha_hora_actual = datetime.datetime.now()
        fecha_hora_actual_formateada = fecha_hora_actual.strftime(FORMATO_FECHA_INVERT)
        fecha_hora_actual_formateada = datetime.datetime.strptime(fecha_hora_actual_formateada, FORMATO_FECHA_INVERT)

        #Valida vencimiento
        #if fecha_hora_actual_formateada > vencimiento:  return (False, 'Cuenta Vencida'), no_cuenta , no_caso , '' 

        sleep(2)
        it.send("{TAB 16}")
        comentarios  = my_copy(driver)

        print(no_cuenta, no_caso, comentarios)


        print('Fin Mis actividades')
        return  (True, ''), no_cuenta, no_caso, comentarios
    
    except Exception as e:
        print(f"No se pudo entrar y buscacar 'Mis Actividades'")
        description_error('08','pantalla_unica_consulta.',e)
        return False, '', ''

def delDownloads():
    for nombre in os.listdir(r'C:\Users\AjusteSucursales\Downloads'):
        ruta_completa = os.path.join(r'C:\Users\AjusteSucursales\Downloads', nombre)
        try:
            if os.path.isfile(ruta_completa) or os.path.islink(ruta_completa):
                os.unlink(ruta_completa)
            elif os.path.isdir(ruta_completa):
                shutil.rmtree(ruta_completa)
        except: pass

def pantalla_caso_negocio(driver, no_cuenta, no_caso): 
    '''
    Funcion que entra en la pantalla de Casos de Negocio y busca Motivo del cliente (Tipo de ajuste)
    '''

    try:
        wait = WebDriverWait(driver, 120)
        act = webdriver.ActionChains(driver)
        delDownloads()
        text_box('Pantalla Casos de Negocio', '☼')

        #Entra a la ventana de Pantalla Unica de Consulta
        open_item_selenium_wait(driver, xpath = home['pantalla_casos']['xpath'] )

        #Espera a que la página tenga el titulo de Home 
        element = wait.until(EC.title_contains('Ordenes'))
        print('Entró a la Pantlla  Casos de Negocio')
        
        #Ingresa Caso de Negocio
        status_open = open_item_selenium_wait(driver ,name = pantalla_casos_negocio['lupa']['name'], xpath =  pantalla_casos_negocio['lupa']['xpath'])
        print(f"Va a buscar el caso: {no_caso}")
        if status_open == False:
            description_error('12','pantalla_casos_negocio','No se encontró la lupa de búsqueda')
            return False
        
        #Se posiciona para escirbir en el Caso de Negocio
        element = driver.find_element(By.XPATH, pantalla_casos_negocio['ingresar_caso']['xpath'])
        element.clear()                    #Limpia lo que haya en el campo
        element.send_keys(no_caso)         #Introduce el numero de orden
        element.send_keys(Keys.RETURN)     #Enter

        sleep(5)

        #Resvisa el motivo del cliente
        print('Motivo del cliente')
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[26]').click()
        motivo_cliente = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[26]/input')
        motivo_cliente = motivo_cliente.get_attribute("value")

        print('Estado original CN')
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[4]').click()
        estadoCN = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[4]/input')
        estadoCN = estadoCN.get_attribute("value")

        print(f'→ Estado inicial CN: {estadoCN}')
        # if 'Cerrado' in estadoCN: return True, 'Cerrado'
        # elif 'Cancelado' in estadoCN: return True, 'Cancelado'
        

        return True, motivo_cliente
    except Exception as e:
        print(f'ERROR en Pantalla Casos de Negocio')
        description_error('11','pantalla_casos_negocio',e)
        return False, ''

def busquedaCol(driver, colBusqueda, path):
    buscandoColumna = True
    contador = 1
    while buscandoColumna == True:
        pathF = path.replace('{contador}', str(contador))
        columna = driver.find_element(By.XPATH, pathF)
        columna = columna.text
        sleep(2)
        print('columan: ', columna)
        if colBusqueda in columna:
            buscandoColumna = False
            return str(contador)
        else:
            contador += 1

def busquedaOculta(driver, xpath, busqueda):

    buscandoCol = True
    contador = 0
    
    while buscandoCol == True:
    
        try:
            contador += 1
            xpathCol = xpath.replace('{contador}', str(contador))
            nameColumna = driver.find_element(By.XPATH, xpathCol)
            columna = driver.execute_script("return arguments[0].textContent;", nameColumna)

            if busqueda in columna:
                print(columna)
                buscandoCol = False
                return str(contador)
            else:
                if contador == 30:
                    error = 'Error Buscando Columna'
                    print(error)
                    buscandoCol = False
                    return error
        
        except:
            if contador == 20:
                error = 'Error Buscando Columna'
                print(error)
                buscandoCol = False
                return error

def validacion_cuenta_cargo_extemporaneo(driver, no_cuenta, no_caso):
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
        wait = WebDriverWait(driver, 150)
        act = webdriver.ActionChains(driver)
        text_box('AJSUTE POR CARGO EXTEMPORANEO', '▬')
        status_pantalla_unica = pantalla_unica_consulta(driver, no_cuenta)
        if status_pantalla_unica == False:
            text_box('Cuenta no valida', '▬')
            error = 'Cuenta Invalida'
            print('Error ', error)
            return False, True , error
        #
        print('▬ Inician validaciones de cuenta ▬')
        sleep(10)


        tipo_cuenta = driver.find_element(By.XPATH, "//input[@aria-label='Tipo']")
        tipo_cuenta = tipo_cuenta.get_attribute("value")
        # tipo_cuenta = driver.execute_script("return arguments[0].textContent;", tipo_cuenta)
        
        print('Tipo Cuenta: ', tipo_cuenta)
        if tipo_cuenta.upper() not in ['RESIDENCIAL', 'NEGOCIO']:
             error = 'Tipo Cuenta ' + tipo_cuenta
             print('Error ', error)
             return False, True, error

        # # 

        # try:
        #      tipo_cuenta = driver.find_element(By.XPATH,pantalla_unica['tipo_cuenta']['xpath'])
        #      tipo_cuenta = tipo_cuenta.get_attribute("value")
        # except Exception:
        #      tipo_cuenta = driver.find_element(By.XPATH,pantalla_unica['tipo_cuenta']['xpath2'])
        #      tipo_cuenta = tipo_cuenta.get_attribute("value")

        
        subtipo_cuenta = driver.find_element(By.XPATH, "//input[@aria-label='SubTipo']")
        subtipo_cuenta = subtipo_cuenta.get_attribute("value")
        # subtipo_cuenta = driver.execute_script("return arguments[0].textContent;", subtipo_cuenta)
        
        print('Subtipo Cuenta', subtipo_cuenta)
        if subtipo_cuenta.upper() not in ['NORMAL']:
            error = 'Subtipo Cuenta ' + subtipo_cuenta
            print('Error ', error)
            return False, True, error

        # try:
        #     subtipo_cuenta = driver.find_element(By.XPATH,pantalla_unica['subtipo_cuenta']['xpath'])
        #     subtipo_cuenta = subtipo_cuenta.get_attribute("value")
        # except Exception:
        #     subtipo_cuenta = driver.find_element(By.XPATH,pantalla_unica['subtipo_cuenta']['xpath2'])
        #     subtipo_cuenta = subtipo_cuenta.get_attribute("value")
        # subtipo_cuenta = my_copy_by_xpath(driver, pantalla_unica['subtipo_cuenta']['xpath'])
        

        sleep(2)
        try:

            driver.find_element(By.XPATH, "//button[@aria-label='Solicitudes de Ajuste Applet de lista:Consulta']").click()
            sleep(5)

            fechaActual = datetime.datetime.now().date()
            fechaActual = fechaActual.replace(day=1)
            fechaActual = fechaActual.strftime('%d/%m/%Y')

            comenntarioFecha = ">= '" + fechaActual + "'"
            print(comenntarioFecha)
            sleep(10)
            columna = busquedaOculta(driver, columnaSA, 'Fecha del ajuste')
            driver.find_element(By.XPATH, filasSA.replace('{contador}', columna)).click()
            sleep(1)
            fechaAjuste = driver.find_element(By.XPATH, filasSA.replace('{contador}', columna) + '/input')
            print('Obteniendo input fecha')
            fechaAjuste.send_keys(str(comenntarioFecha))
            sleep(5)

            columna = busquedaOculta(driver, columnaSA, 'Estado')
            driver.find_element(By.XPATH, filasSA.replace('{contador}', columna)).click()
            sleep(1)
            estadoAjuste = driver.find_element(By.XPATH, filasSA.replace('{contador}', columna) + '/input')
            print('Obteniendo input fecha')
            estadoAjuste.send_keys('Aplicado')
            estadoAjuste.send_keys(Keys.RETURN)
            sleep(5)

            columna = busquedaOculta(driver, columnaSA, 'Motivo del ajuste')
            driver.find_element(By.XPATH, filasSA.replace('{contador}', columna)).click()
            sleep(1)
            motivoAjuste = driver.find_element(By.XPATH, filasSA.replace('{contador}', columna) + '/input')
            motivoAjuste.send_keys('CARGO POR PAGO EXTEMPORANEO')
            motivoAjuste.send_keys(Keys.RETURN)
            motivoAjuste.send_keys(Keys.RETURN)

        except Exception as e:
            print('Falla al filtrar los ajustes previos')
            print(e)

        sleep(8)

        try:
            columna = busquedaOculta(driver, columnaSA, 'Fecha del ajuste')
            driver.find_element(By.XPATH, filasSA.replace('{contador}', columna)).click()
            error = 'Ajuste Mes'
            print('###########################')
            print(error)
            print('###########################')
            print('Validaciones Completa c:')
            return False, True, error
        except:
            print('Sin ajuste en el mes')
        
        sleep(4)

        try:
            driver.find_element(By.XPATH, "//button[@aria-label='Solicitudes de Ajuste Applet de lista:Consulta']").click()
            sleep(5)

            fechaActual = datetime.date.today()
            fechaActual = fechaActual.replace(day=1)
            mesAnterior = fechaActual - datetime.timedelta(days=1)
            mesAnterior = mesAnterior.strftime('%d/%m/%Y')

            fecha6MesesP = datetime.date.today()
            fecha6MesesP = fecha6MesesP - datetime.timedelta(days=30*3)
            fecha6MesesP = fecha6MesesP.replace(day=1)
            fecha6MesesP = fecha6MesesP.strftime('%d/%m/%Y')

            comenntarioFecha = ">= '{}' AND <= '{}'".format(fecha6MesesP, mesAnterior)
            print(comenntarioFecha)
            sleep(10)
            
            columna = busquedaOculta(driver, columnaSA, 'Fecha del ajuste')
            driver.find_element(By.XPATH, filasSA.replace('{contador}', columna)).click()
            sleep(1)
            fechaAjuste = driver.find_element(By.XPATH, filasSA.replace('{contador}', columna) + '/input')
            print('Obteniendo input fecha')
            fechaAjuste.send_keys(str(comenntarioFecha))
            sleep(5)

            columna = busquedaOculta(driver, columnaSA, 'Motivo del ajuste')
            driver.find_element(By.XPATH, filasSA.replace('{contador}', columna)).click()
            sleep(1)
            motivoAjuste = driver.find_element(By.XPATH, filasSA.replace('{contador}', columna) + '/input')
            motivoAjuste.send_keys('CARGO POR PAGO EXTEMPORANEO')
            motivoAjuste.send_keys(Keys.RETURN)
            motivoAjuste.send_keys(Keys.RETURN)
        except Exception as e:
            print('Falla al filtrar los ajustes previos')
            print(e)

        sleep(8)

        try:
            columna = busquedaOculta(driver, columnaSA, 'Fecha del ajuste')
            driver.find_element(By.XPATH, filasSA.replace('{contador}', columna)).click()
            error = 'Ajuste Reciente'
            print(error)
            print('Validaciones Completa c:')
            return False, True, error

            
        except Exception:
            print('Sin ajustes previos')
            sleep(7)
        
        driver.find_element(By.XPATH, "//button[@aria-label='Solicitudes de Ajuste Applet de lista:Consulta']").click()
        sleep(5)
        columna = busquedaOculta(driver, columnaSA, 'Fecha del ajuste')
        driver.find_element(By.XPATH, filasSA.replace('{contador}', columna)).click()
        sleep(1)
        fechaAjuste = driver.find_element(By.XPATH, filasSA.replace('{contador}', columna) + '/input')
        sleep(1)
        fechaAjuste.send_keys('')
        fechaAjuste.send_keys(Keys.RETURN)


        sleep(1)
        driver.find_element(By.XPATH, "//button[@aria-label='Casos de Negocio Applet de lista:Consulta']").click()
        sleep(5)
        columna = busquedaOculta(driver, columnaCN, 'Motivo Cliente')
        driver.find_element(By.XPATH, filasCN.replace('{contador}', columna)).click()
        columna = driver.find_element(By.XPATH, filasCN.replace('{contador}', columna) + "/input")
        columna.send_keys('CARGO EXTEMPORANEO')
        sleep(1)
        columna.send_keys(Keys.RETURN)
        columna.send_keys(Keys.RETURN)
        sleep(8)
        try: 
            columna = busquedaOculta(driver, columnaCN, 'Motivo Cliente')
            driver.find_element(By.XPATH, filasCN.replace('{contador}', columna)).click()
        except:
            error = 'Sin CN Sucursal'
            print('Error ', error)
            return False, True, error

        
        print('Validaciones Completa c:')
        return True, '',''



    except Exception as e:
        print(f'ERROR ajustando el cargo extemporaneo. Caso NO. {no_caso}. CUENTA: {no_cuenta}')
        description_error('11','validacion_cuenta_cargo_extemporaneo',e)
        print('Error ', e)
        return False, True, 'Excepcion Validacion'

def cargandoElemento(driver, elemento, atributo, valorAtributo, path = False):

    cargando = True
    contador = 0

    while cargando:

        sleep(1)
        try: 
            
            if path == False: 
                driver.find_element(By.XPATH, f"//{elemento}[@{atributo}='{valorAtributo}']").click()
                return True
            else: 
                print('aqui')
                driver.find_element(By.XPATH, path).click()
                return True
            
        
        except:
            contador += 1
            if contador == 60: return False


######################### Funcion Nueva
def obtencionColumna(driver, nombreColumna, path, path2 = False):

    buscandoColumna = True
    contador = 0

    while buscandoColumna:

        try:
            contador += 1
            nameColumna2 = 'False'
            pathF = path.replace('{contador}', str(contador))

            try:
                nameColumna = driver.find_element(By.XPATH, pathF)
                nameColumna = driver.execute_script("return arguments[0].textContent;", nameColumna)
                print(f'path1: {nameColumna}')
            except: nameColumna = 'False'

            if path2 != False: 
                try:
                    pathF2 = path2.replace('{contador}', str(contador))
                    nameColumna2 = driver.find_element(By.XPATH, pathF2)
                    nameColumna2 = driver.execute_script("return arguments[0].textContent;", nameColumna2)
                except: nameColumna2 = 'False'
                print(f'path2: {nameColumna2}')

            if nombreColumna in nameColumna or nombreColumna in nameColumna2: return str(contador)
            else:
                if contador == 100: return False

        except Exception as e: print(str(e)); return False

######################### Funcion Nueva
def extraer_texto_de_pdf_con_imagenes(ruta_pdf):
    try:
        print('Iniciando LEctura PDF')
        print(ruta_pdf)
        # Instancia de tesseract
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        print(1)
        # Instancia de Poppler
        poppler_path = r'C:\Program Files\poppler-24.07.0\Library\bin'
        print(2)

        # Convertir PDF a imagenes
        imagenes = convert_from_path(ruta_pdf, poppler_path=poppler_path)
        print(3)

        # Obtener la ultima pagina del PDF
        imagen = imagenes[len(imagenes)-1]
        print(4)

        # Inicio de extraccion de texto

        texto_completo = ""
        print(5)
        temp_image = f"temp_image.png"
        print(6)
        imagen.save(temp_image, 'PNG')
        print(7)

        # OCR en la imagen
        texto = pytesseract.image_to_string(Image.open(temp_image))
        print(8)
        texto_completo += f"Página 1:\n{texto}\n\n"
        print(9)

        # Remove de la imagen tmp
        os.remove(temp_image)

        if "CARGO POR PAGO TARDIO $" in texto_completo.upper():
            print('Cargo detectado')

            texto = texto_completo.upper().split("CARGO POR PAGO TARDIO $", 1)
            if len(texto) > 1:
                importeAjuste = texto[1]
                importeAjuste = importeAjuste.splitlines()[0].strip()
                print(f'Monto del pago tardío: {importeAjuste}')
                return True, importeAjuste
            else:
                print('Monto de Pago Tardío no detectado')
                return False, 'Monto Pago No Detectado'

        else: print('NO'); return False, 'Sin Pago Tardio'

    except Exception as e: print(e); sleep(10000)




def busqueda_factura(driver, no_cuenta, no_caso):
    '''
    Funcion que busca y valida que la factura sean correctas
    args:
        - driver
        - no_cuenta
        - no_caso

    out:
        - Bool: True n caso de localizar la factura y que en su descripción venga la leyenda:
            "Cargo por pago tardío".
        - fecha_factura: regresa el mes de la factura
    ''' 

    pathEncabezadoFacturas = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
    pathFacturas = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[{contador2}]/td[{contador}]'

    pathCuadroPDF = '/html/body/div[21]/div[2]'
    pathEmbedPDF = '/html/body/div[21]/div[2]/embed'
    pathCierrePDF = '/html/body/div[21]/div[1]/button'
    
    try:

        # try: driver.find_element(By.XPATH, "//a[@title='Pantalla Única de Consulta']").click()
        # except: return 'Registro Pendiente', '', '-', '-'
        
        # # Buscando Elemento
        # lupa_busqueda_cuenta = cargandoElemento(driver, 'button', 'title', 'Pantalla Única de Consulta Applet de formulario:Consulta')
        # if lupa_busqueda_cuenta == False: return 'Inconsistencia Siebel: Pantalla NO Carga', '-', '-', '-'
        # sleep(5)

        # # Ingreso Cuenta
        # input_busqueda_cuenta = driver.find_element(By.XPATH, "//input[@aria-label='Numero Cuenta']")
        # input_busqueda_cuenta.click()
        # sleep(1)
        # input_busqueda_cuenta.send_keys(no_cuenta)
        # input_busqueda_cuenta.send_keys(Keys.RETURN)
        # print('♦ Cuenta Ingresada ♦')

        # lupa_busqueda_cuenta = cargandoElemento(driver, 'span', 'id', 'Saldo_Vencido_Label_12')
        # if lupa_busqueda_cuenta == False: return 'Inconsistencia Siebel: Pantalla NO Carga', '-', '-', '-'
        # print('♥ Cuenta OK! ♥')
        # sleep(5)

        text_box('BUSQUEDA DE FACTURA', '▬')

        # Obtencion posicion campo Fecha Limite de la Factura

        posFechaLimite = obtencionColumna(driver, 'Fecha límite', pathEncabezadoFacturas)
        
        # Busqueda y descarga de Facturas

        fechaHoy = datetime.datetime.today()
        rango3Meses = fechaHoy - datetime.timedelta(days=91)
        buscandoFacturas = True
        facturasDescargadas = []
        fechasValidadas = []
        contador = 2
        while buscandoFacturas:
            try:

                # Obtencion Fecha de Factura
                fechaFactura = driver.find_element(By.XPATH, pathFacturas.replace('{contador2}', str(contador)).replace('{contador}', str(posFechaLimite)))
                fechaFactura = fechaFactura.get_attribute('title')
                print(fechaFactura)

                # Validacion Fecha Unica
                if fechaFactura not in fechasValidadas: 
                    fechasValidadas.append(fechaFactura)
                    fechaFactura = datetime.datetime.strptime(fechaFactura, '%d/%m/%Y')

                    # Validacion rango 3 meses fecha obtenida
                    if rango3Meses <= fechaFactura <= fechaHoy or fechaFactura >= fechaHoy:
                        driver.find_element(By.XPATH, pathFacturas.replace('{contador2}', str(contador)).replace('{contador}', str(posFechaLimite))).click()
                        sleep(1)
                        print('Esperando Cuadro Factura')

                        driver.find_element(By.XPATH, "//button[@title='Ver Factura']").click()

                        # Inicio de proceso de descarga de factura
                        cargandoFactura = True
                        while cargandoFactura:
                            # try: 
                            #     alert = Alert(driver)
                            #     alert_txt = alert.text
                                
                            try:
                                print('Descargando Factura')
                                # Cuadro Emergente con la factura
                                driver.find_element(By.XPATH, pathCuadroPDF)
                                sleep(7)

                                # Obtencion del url para descargar factura
                                url = driver.find_element(By.XPATH, pathEmbedPDF)
                                url = url.get_attribute('src')
                                print(url)
                                
                                if 'blob' in url:

                                    # Inyeccion de codigo para descarga de pdf
                                    driver.execute_script("""
                                        // arguments[0] es el blob que se pasa
                                        const blobUrl = arguments[0];
                                        const a = document.createElement('a');
                                        a.style.display = 'none';
                                        a.href = blobUrl;
                                        a.download = 'documento.pdf';
                                        document.body.appendChild(a);
                                        a.click()
                                        document.body.removeChild(a);
                                    """, url)

                                    # Inicio de validacion de pdf unico descargado
                                    descargandoPDF = True
                                    contadorDescarga = 0
                                    while descargandoPDF:

                                        ruta = Path(r'C:\Users\AjusteSucursales\Downloads')
                                        pdfs = list(ruta.glob('*.pdf'))
                                        if not pdfs:
                                            contadorDescarga += 1
                                            sleep(1)
                                            if contadorDescarga == 20: return False, None, 'Inconsistencia Siebel: PDF No descarga', '-'
                                        else:
                                            reciente = max(pdfs, key=lambda f: f.stat().st_mtime)
                                            if reciente not in facturasDescargadas:
                                                print('PDF MAS RECIENTE')
                                                facturasDescargadas.append(reciente)
                                                resultado, monto = extraer_texto_de_pdf_con_imagenes(reciente)
                                                if resultado == True:
                                                    descargandoPDF = False
                                                    cargandoFactura = False
                                                    buscandoFacturas = False
                                                    fechaFactura = fechaFactura.strftime('%d/%m/%Y')
                                                    print('→ Fin de extraccion de facturas')
                                                    try: driver.find_element(By.XPATH, pathCierrePDF).click()
                                                    except: pass
                                                    return True, fechaFactura,'', monto
                                                # elif len(facturasDescargadas) == 3 and resultado == False: return False, None, 'No Aplica: Item Facturacion', '-'
                                                else:
                                                    try: driver.find_element(By.XPATH, pathCierrePDF).click()
                                                    except: pass
                                                    sleep(5)
                                                    descargandoPDF = False
                                                    cargandoFactura = False
                                                    contador += 1
                                            else:
                                                sleep(1)
                                                contadorDescarga += 1
                                                if contadorDescarga == 20: 
                                                    descargandoPDF = False
                                                    cargandoFactura = False
                                                    contador += 1
                                
                                try: driver.find_element(By.XPATH, pathCierrePDF).click()
                                except: pass

                                sleep(7)
                            except: 
                                sleep(1)
                                print('Esperando ventana PDF')
                    # elif fechaFactura > fechaHoy: contador += 1
                    else:
                        print('Fin de busqueda de Facturas')
                        buscandoFacturas = False
                        return False, None, 'No Aplica: Item Facturacion', '-'

                
                else: contador += 1

            except Exception as e:
                try: 
                    driver.find_element(By.XPATH, "//td[@id='next_pager_s_3_l']").click()
                    contador = 2
                    sleep(7)
                except: 
                    buscandoFacturas = False
                    return False, None, 'No Aplica: Item Facturacion', '-'
        
        # sleep(10000)


        # #Historial de Facturas
        # print('Entra Historial de Facturas')
        # if open_item_selenium_wait(driver, xpath =  pantalla_unica['historial_facturas']['xpath'] ) != True: return False, None, 'Falla Historial Facturas', ''
        # sleep(3)

        # print('Factura Actual')
        # if open_item_selenium_wait(driver, xpath =  pantalla_unica['factura_actual']['xpath'] )!= True: return False, None, 'Falla Factura Actual', ''
        # #Identifica periodos, esto hasta deberia ser una funcion
        # sleep(5)
        # it.send('{UP 50}')
        # print('Scroll hacia arriba') 
        # sleep(2)
        # print('Recorre los items')
        # #Resvia las ultimas 6 facturas
        # busquedaFactura = True
        # contadorBusqueda= 0
        # contadorFactura = 1
        # while busquedaFactura == True:
        #     try:
        #         #Valida fecha copiada
        #         print('Factura #',str(contadorFactura))
                
        #         driver.find_element(By.XPATH, f'//*[@id="{str(contadorFactura)}_s_1_l_PeriodStartDate"]').click()
        #         fecha_fatcura = driver.find_element(By.XPATH, f'//*[@id="{str(contadorFactura)}_PeriodStartDate"]')
        #         fecha_fatcura = fecha_fatcura.get_attribute("value")
        #         fecha_fatcura = fecha_fatcura.split('/')

        #         status_descripcion, importe = busqueda_descripcion(driver, "Cargo por pago tardío" )
        #         if  status_descripcion == True: return True, obtener_nombre_mes(fecha_fatcura[1]),'', importe
        #         else: 
        #             contadorBusqueda += 1
        #             contadorFactura += 1
                    
        #             # if contadorBusqueda == 16:
        #             #     busquedaFactura = False
        #             #     return False, None, 'Item Facturacion', ''
        #             # else:
        #             #     contadorFactura += 1
        #     except Exception as e:
        #         print(e)
        #         # sleep(1000)
        #         try:
        #             driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[5]/div/table/tbody/tr/td[2]/table/tbody/tr/td[3]/span').click()
        #             contadorFactura = 1
        #             sleep(2)
        #         except Exception as e:
        #             print(e)
        #             print('################ FIN BUSQUEDA FACTURA ###################')
        #             sleep(2)
        #             return False, None, 'Item Facturacion', ''
        
        # return False, None, 'Item Facturacion', ''

    except Exception as e:
        print(f'ERROR ajustando el cargo extemporaneo. Caso NO. {no_caso}. CUENTA: {no_cuenta}')
        description_error('12','busqueda_factura',e)
        error = e
        print('Error ',error)
        return False, '',error

def busqueda_descripcion(driver,descripcion ):
    '''
    Funcion que busca la descripcion en las facturas seleccionadas
    arg:
        - driver: obj
        - descripccion: str

    out:
        - Bool
    '''  
    sleep(2)
    open_item_selenium_wait(driver, xpath =  pantalla_unica['detalles']['xpath'] )

    i = 1
    busquedaFactura = True
    while busquedaFactura == True:
        try:
            driver.find_element(By.XPATH, f'//*[@id="{str(i)}_s_2_l_Comments"]').click()
            sleep(2)
            facturas = driver.find_element(By.XPATH, f'//*[@id="{str(i)}_Comments"]')
            facturas = facturas.get_attribute("value")

            if descripcion in facturas:
                driver.find_element(By.XPATH, f'//*[@id="{str(i)}_s_2_l_Amount"]').click()
                sleep(2)
                importe = driver.find_element(By.XPATH, f'//*[@id="{str(i)}_Amount"]')
                importe = importe.get_attribute("value")
                print('Factura Encontrada')
                print('♦ Descripcion localizado ♦')
                print('Regresa a Pantalla Unica ')
                sleep(2)
                it.send('{UP 20}')
                open_item_selenium_wait(driver, xpath ='/html/body/div[1]/div/div[5]/div/div[8]/div[1]/div/div[3]/ul/li[1]/span/a')
                return True, importe

            else:
                i += 1
                print('Descripcion no encontrada')
        except Exception as e:
            print(e)
            # sleep(1000)
            try:
                driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div/form/span/div/div[3]/div/div/div[5]/div/table/tbody/tr/td[2]/table/tbody/tr/td[3]/span').click()
                i = 1
                sleep(2)
            except Exception as e:
                return False, ''
                
        
    

    # try:
    #     for i in range(1,8):

    #         driver.find_element(By.XPATH, f'//*[@id="{str(i)}_s_2_l_Comments"]').click()
    #         sleep(2)
    #         facturas = driver.find_element(By.XPATH, f'//*[@id="{str(i)}_Comments"]')
    #         facturas = facturas.get_attribute("value")

    #         if descripcion in facturas:
    #             driver.find_element(By.XPATH, f'//*[@id="{str(i)}_s_2_l_Amount"]').click()
    #             sleep(2)
    #             importe = driver.find_element(By.XPATH, f'//*[@id="{str(i)}_Amount"]')
    #             importe = importe.get_attribute("value")
    #             print('Factura Encontrada')
    #             print('♦ Descripcion localizado ♦')
    #             print('Regresa a Pantalla Unica ')
    #             sleep(2)
    #             it.send('{UP 20}')
    #             open_item_selenium_wait(driver, xpath ='/html/body/div[1]/div/div[5]/div/div[8]/div[1]/div/div[3]/ul/li[1]/span/a')
    #             return True, importe

    #         else:
    #             print('Descripcion no encontrada')

    #     return False, ''

    # except Exception as e:
    #     print('No hay mas elementos en los detalles')
    #     print(e)
    #     return False, ''
    


def aplicacion_ajuste_cargo_extemporaneo(driver, mes, cn, importe):
    '''
    Funcion que aplica el ajuste
    arg: 
        - driver
    out:
        - bool: True en caso de que se cree el ajuste, se guarde y se envie
    '''
    try:
        text_box('AJUSTE POR CARGO EXTEMPORANEO', '▬')
        print('Cargando...')
        
        cargandoPatalla = True
        contador = 0

        while cargandoPatalla == True:
            try:
                driver.find_element(By.XPATH, "//button[@aria-label='Solicitudes de Ajuste Applet de lista:Nuevo']").click()
                cargandoPatalla = False
            except:
                sleep(10)
                contador += 1
                if contador == 10:
                    cargandoPatalla = False
                    return False, 'excepcio'
        sleep(5)
        importe = importe.replace('$', '')
        element = driver.find_element(By.XPATH,  solicitud_ajuste['importe']['xpath'])
        element.clear()                    #Limpia lo que haya en el campo
        element.send_keys(importe)           #Introduce el numero de orden
        sleep(3)
        element = driver.find_element(By.XPATH, ajuste_CE['aplicar'])
        element.clear()                    #Limpia lo que haya en el campo
        element.send_keys('A favor')
        sleep(3)
        element = driver.find_element(By.XPATH, ajuste_CE['motivo'])
        element.clear()                    #Limpia lo que haya en el campo
        element.send_keys('CARGO POR PAGO EXTEMPORANEO')
        sleep(3)
        element = driver.find_element(By.XPATH, ajuste_CE['comentario'])
        element.clear()                    #Limpia lo que haya en el campo
        element.send_keys(f'''SE APLICA AJUSTE POR CARGO EXTEMPORANEO FACTURA DEL MES DE {mes} CN {cn} ROBOT''')
        sleep(3)
        numAjuste = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/div/table[2]/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/span/input')
        numAjuste = numAjuste.get_attribute("value")
        print(numAjuste)
        sleep(3)
        print('Guardar')
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/div/table[1]/tbody/tr/td[1]/span/div/div[3]/button[1]').click()
        sleep(2)
        driver.find_element(By.XPATH,'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/span/div/div[1]/div[2]/button[2]').click()
        sleep(3)
        numajusteCampo =  driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/input')
        numajusteCampo.send_keys(numAjuste)
        numajusteCampo.send_keys(Keys.RETURN)
        sleep(6)
        print('Enviar')
        try:
            driver.find_element(By.XPATH, "//button[@aria-label='Solicitudes de Ajuste Applet de lista:Enviar']").click()
            sleep(10)
            print('BTN Aceptar envio')

            driver.find_element(By.XPATH, "//button[@aria-label='Enviar Ajuste Applet de formulario:Aceptar']").click()
            sleep(10)
        except:
            error = 'Envio Ajuste'
            return False, error
        sleep(5)
        print('Consulta saldos')
        it.send('{UP 30}')
        open_item_selenium_wait(driver, xpath =  solicitud_ajuste['consulta_saldos']['xpath'])
        sleep(3)
        open_item_selenium_wait(driver ,name = pantalla_unica['saldo_pendiente']['name'], xpath =  pantalla_unica['saldo_pendiente']['xpath'])
        print('AJUSTE REALIZADO CON EXITO ☺')
        
        return True,''
        

    
    except Exception as e:
        print(f'ERROR ajustando el cargo extemporaneo. Caso NO. . CUENTA:')
        description_error('13','aplicacion_ajuste_cargo_extemporaneo',e)
        return False, 'excepcio,'

def busquedaCampo(driver, colBusqueda):
    buscandoColumna = True
    contador = 1
    while buscandoColumna == True:
        columna = driver.find_element(By.XPATH, f"/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{str(contador)}]/div")
        columna = columna.text
        sleep(2)
        print('columan: ', columna)
        if colBusqueda in columna:
            buscandoColumna = False
            return str(contador)
        else:
            contador += 1

def cierre_caso_y_actividad(driver, numero_cuenta, cn, mes, tipoCargo):
    '''
    Funicon que cierra el caso de negocio y la actividad

    args: 
        - driver
        - numero_cuenta
        - caso_negocio
        - mes
    
    out:
        - bool
    '''
    try:
        text_box('CIERRE DE CN Y ACTIVIDAD', '▬')

        #Ingresa Caso de Negocio
        
        sleep(20)
        driver.find_element(By.XPATH, "//button[@aria-label='Casos de Negocio Applet de lista:Consulta']").click()
        sleep(5)

        columna = busquedaCol(driver, 'Motivo', columnaCN)
        driver.find_element(By.XPATH, filasCN.replace('{contador}', columna)).click()

        columna = busquedaCol(driver, 'Caso de Negocio', columnaCN)
        driver.find_element(By.XPATH, filasCN.replace('{contador}', columna)).click()
        casoNegocio = driver.find_element(By.XPATH, filasCN.replace('{contador}', columna) + "/input")
        casoNegocio.clear()                         # Limpia lo que haya en el campo
        casoNegocio.send_keys(cn)         # Introduce el numero de orden
        casoNegocio.send_keys(Keys.RETURN)
        sleep(8)
        columna = busquedaCol(driver, 'Motivo', columnaCN)
        driver.find_element(By.XPATH, filasCN.replace('{contador}', columna)).click()

        print('Profundiza en el caso de negocio')
        driver.find_element(By.XPATH, profundizarCN).click()
        sleep(8)

        print('Ingresa en las actividades')
        driver.find_element(By.XPATH, caso_negocio['actividades']).click()
        sleep(6)

        print('Busca los comentarios')
        columna = busquedaOculta(driver, caso_negocio['columnasAct'], 'Comentarios')
        driver.find_element(By.XPATH, caso_negocio['filasAct'].replace('{contador}', columna)).click()
        sleep(1)
        comentario = driver.find_element(By.XPATH, caso_negocio['filasAct'].replace('{contador}', columna) + '/textarea')
        comentario.clear()
        if tipoCargo == 'Extemporaneo':
            comentario.send_keys(f'''SE APLICA AJUSTE POR CARGO EXTEMPORANEO FACTURA DEL MES DE {mes} CN {cn} ROBOT''')
        elif tipoCargo == 'Cobranza':
            comentario.send_keys(f'''SE APLICA AJUSTE POR CONVENIO COBRANZA ROBOT''')
        elif tipoCargo == 'Otro':
            comentario.send_keys(f'''SE APLICA AJUSTE POR CONVENIO COBRANZA ROBOT''')
        elif tipoCargo == 'Otro':
            comentario.send_keys(f'''SE APLICA AJUSTE POR CONVENIO COBRANZA ROBOT''')
        sleep(3)

        #------------
        print('Ingresa Motivo de la Cancelacion') #AJUSTAR SEGUN EL TIPO
        columna = busquedaOculta(driver, caso_negocio['columnasAct'], 'Motivo del cierre')
        driver.find_element(By.XPATH, caso_negocio['filasAct'].replace('{contador}', columna)).click()
        sleep(1)
        motivo = driver.find_element(By.XPATH, caso_negocio['filasAct'].replace('{contador}', columna) + '/input')
        motivo.clear()
        motivo.send_keys('SE APLICA AJUSTE')
        sleep(2)

        #------------
        print('Ingresa Estado')
        columna = busquedaOculta(driver, caso_negocio['columnasAct'], 'Estado')
        driver.find_element(By.XPATH, caso_negocio['filasAct'].replace('{contador}', columna)).click()
        sleep(2)
        driver.find_element(By.XPATH, caso_negocio['selecEstado'].replace('{contador}', columna)).click()
        sleep(3)
        columna = busquedaOculta(driver, caso_negocio['estado'], 'CERRADA')
        driver.find_element(By.XPATH, caso_negocio['estado'].replace('{contador}', columna)).click()
        sleep(3)

        # it.send("{CTRLDOWN}S{CTRLUP}") #SE COMENTA ESTA LINEA PARA PRUEBAS
        driver.find_element(By.XPATH, "//button[@aria-label='Caso de negocio Applet de formulario:Guardar']").click()
        sleep(5)
        try:
            
            alert = Alert(driver)
            alert_txt = alert.text
            print(alert_txt)
            # sleep(1000)
            alert.accept()
        
        except Exception:

            print('Envio Correcto')

        driver.find_element(By.XPATH, '//a[contains(text(), "Más información")]').click()
        sleep(15)
        print('Estado original CN')
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[4]').click()
        estadoCN = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[4]/input')
        estadoCN = estadoCN.get_attribute("value")

        print(f'→ Estado actualizado CN: {estadoCN}')
        if 'Cerrado' in estadoCN: return True, 'Cerrado'
        else: return True, 'Pendiente'


        # print('Verificando el status del CN')                                                                                 
        # statusCN = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/div/span/div[3]/div/div/table/tbody/tr[3]/td[8]/div/input").click()
        # sleep(2)
        # print('Obteniendo Valor')
        # statusCN = statusCN.get_attribute("value")
        # print(statusCN)
        # sleep(2)
        # if 'Abierto' in statusCN:
        #     return False, 'Pendiente'
        


        return True, ''
    except Exception as e:
        print(f'ERROR CERRANDO EL CASO Y LA ACTIVDIAD. Caso NO. {caso_negocio} . CUENTA: {numero_cuenta}')
        description_error('14','cierre_caso_y_actividad',e)
        return False, 'Excepcion Cierre CN y Actividad'
    
def cancelar_caso(driver, numero_cuenta, cn, motivo_cancelacion):
    '''
    Funcion que cancela la actividad debido a que no se cumple una o mas de las actividades
    '''
    try:
        text_box('CANCELAR DE CN Y ACTIVIDAD', '▬')
        
        print(print('Ingresa a Pantalla Unica'))
        pantalla_unica_consulta(driver, numero_cuenta)

        
        print('Se va al recuadro de los casos de negocio')

        #Ingresa Caso de Negocio
        sleep(20)
        driver.find_element(By.XPATH, "//button[@aria-label='Casos de Negocio Applet de lista:Consulta']").click()
        sleep(5)
        
        columna = busquedaCol(driver, 'Motivo', columnaCN)
        driver.find_element(By.XPATH, filasCN.replace('{contador}', columna)).click()

        columna = busquedaCol(driver, 'Caso de Negocio', columnaCN)
        driver.find_element(By.XPATH, filasCN.replace('{contador}', columna)).click()
        casoNegocio = driver.find_element(By.XPATH, filasCN.replace('{contador}', columna) + "/input")
        casoNegocio.clear()                         # Limpia lo que haya en el campo
        casoNegocio.send_keys(cn)         # Introduce el numero de orden
        casoNegocio.send_keys(Keys.RETURN)
        sleep(8)
        columna = busquedaCol(driver, 'Motivo', columnaCN)
        driver.find_element(By.XPATH, filasCN.replace('{contador}', columna)).click()

        print('Profundiza en el caso de negocio')
        driver.find_element(By.XPATH, profundizarCN).click()
        sleep(8)

        print('Ingresa en las actividades')
        driver.find_element(By.XPATH, caso_negocio['actividades']).click()
        sleep(6)

        print('Busca los comentarios')
        columna = busquedaOculta(driver, caso_negocio['columnasAct'], 'Comentarios')
        driver.find_element(By.XPATH, caso_negocio['filasAct'].replace('{contador}', columna)).click()
        sleep(1)
        comentario = driver.find_element(By.XPATH, caso_negocio['filasAct'].replace('{contador}', columna) + '/textarea')
        comentario.clear()
        comentario.send_keys(f'''SE CANCELA CN {cn} POR FALTA DE INFORMACION {motivo_cancelacion}. ROBOT''')
        sleep(3)

        #------------
        print('Ingresa Motivo de la Cancelacion') #AJUSTAR SEGUN EL TIPO
        columna = busquedaOculta(driver, caso_negocio['columnasAct'], 'Motivo de la cancelación')
        driver.find_element(By.XPATH, caso_negocio['filasAct'].replace('{contador}', columna)).click()
        sleep(1)
        motivo = driver.find_element(By.XPATH, caso_negocio['filasAct'].replace('{contador}', columna) + '/input')
        motivo.clear()
        motivo.send_keys(motivo_cancelacion)
        sleep(2)

        #------------
        print('Ingresa Estado')
        columna = busquedaOculta(driver, caso_negocio['columnasAct'], 'Estado')
        driver.find_element(By.XPATH, caso_negocio['filasAct'].replace('{contador}', columna)).click()
        sleep(2)
        driver.find_element(By.XPATH, caso_negocio['selecEstado'].replace('{contador}', columna)).click()
        sleep(3)
        columna = busquedaOculta(driver, caso_negocio['estado'], 'CANCELADA')
        driver.find_element(By.XPATH, caso_negocio['estado'].replace('{contador}', columna)).click()
        sleep(3)

        # it.send("{CTRLDOWN}S{CTRLUP}") #SE COMENTA ESTA LINEA PARA PRUEBAS
        driver.find_element(By.XPATH, "//button[@aria-label='Caso de negocio Applet de formulario:Guardar']").click()
        sleep(5)

        # print('Verificando el status del CN')
        # statusCN = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/div/span/div[3]/div/div/table/tbody/tr[3]/td[8]/div/input").click()
        # sleep(2)
        # print('obteninedo el valor')
        # statusCN = statusCN.get_attribute("value")
        # print(statusCN)
        # sleep(2)
        # if 'Abierto' in statusCN:
        #     return False, 'Pendiente'

        try:
            
            alert = Alert(driver)
            alert_txt = alert.text
            print(alert_txt)
            alert.accept()
        
        except Exception:

            print('Envio Correcto')

        driver.find_element(By.XPATH, '//a[contains(text(), "Más información")]').click()
        sleep(15)
        print('Estado original CN')
        driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[4]').click()
        estadoCN = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[4]/input')
        estadoCN = estadoCN.get_attribute("value")

        print(f'→ Estado actualizado CN: {estadoCN}')
        if 'Cancelado' in estadoCN: return 'Cancelado'
        else: return 'Pendiente'


        # print('GUARDA LOS CAMBIOS') 
        # it.send('+{TAB 5}')
        # it.send("{CTRLDOWN}S{CTRLUP}") #SE COMENTA ESTA LINEA PARA PRUEBAS
        #---

        # print('Valida que se haya aplicado')
        # it.send('{UP 20}')
        # sleep(2)
        # open_item_selenium_wait(driver, name = 's_1_1_141_0', xpath = '//*[@id="a_1"]/div/table/tbody/tr[3]/td[8]/div')
        # estado_caso = my_copy(driver)
        # if estado_caso == 'Cancelado':
        #     print('Caso Cancelado de forma exitosa')

        # return True #Debee estar en el if

    except Exception as e:

        print(f'ERROR CERRANDO EL CASO Y LA ACTIVDIAD. Caso NO. {caso_negocio} . CUENTA: {numero_cuenta}')
        description_error('14','cierre_caso_y_actividad',e)
        return False

def obtener_nombre_mes(numero_mes):
    meses = {
        1: "Enero",
        2: "Febrero",
        3: "Marzo",
        4: "Abril",
        5: "Mayo",
        6: "Junio",
        7: "Julio",
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre"
    }
    nombre_mes = meses.get(int(numero_mes))
    if nombre_mes:
        return nombre_mes
    else:
        return "Número de mes inválido"
    

# user = "apiliado"
# password = 'Botiendas.2023'
# driver, status_logueo = login_siebel(user, password)
# cuenta = "35752303"
# CN = "1149317453910"
# validacion_cuenta_cargo_extemporaneo(driver, cuenta, CN )  
# busqueda_factura(driver, cuenta, CN )  
# aplicacion_ajuste_cargo_extemporaneo(driver, 'enero', CN )                                       
# cierre_caso_y_actividad(driver,  cuenta, CN , 'ENERO')