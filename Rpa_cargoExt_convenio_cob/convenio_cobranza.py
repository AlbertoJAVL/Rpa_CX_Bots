#----------Selenium--------------------#
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.alert import Alert
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

#---------Variables globales---------------#

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
        inputOrdenarPor.clear()
        sleep(3)
        inputOrdenarPor.send_keys(elementoOrden)
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

def ordenar2(driver, engrane, opcionOrden, input, elementoOrden, orden, confirmar):

    ########################################################################

                ### FUNCION PARA ORDENAR TABLAS ###

    ########################################################################
    try:
        print('Ordenamiento de tabla')

        driver.find_element(By.XPATH, engrane).click() #Click sobre engrane
        sleep(3)
        driver.find_element(By.XPATH, opcionOrden).click() #Click sobre opcion "Ordenar"
        sleep(8)
        it.send('{TAB 6}')
        sleep(2)
        it.send('{ENTER}')
       

        print('Tabla ordenada')
    except Exception as e:
        print(e)
        print('------------- Fallo al ordenar la tabla ---------------')
        sleep(1000)

def buscar(driver, lupa, campo, busqueda):
    print('Funcion buscar')
    driver.find_element(By.XPATH, lupa).click() #Click sobre lupa de buscar
    sleep(4)

    driver.find_element(By.XPATH, campo).click()

    element = driver.find_element(By.XPATH, campo)
    
    element.clear()                    #Limpia lo que haya en el campo
    element.send_keys(busqueda)         #Introduce el motivo del ajuste
    element.send_keys(Keys.RETURN)     #Enter
    element.send_keys(Keys.RETURN)     #Enter 

def formatoFecha(fecha, time):

    print('Formateando fecha')

    if time == True:
        fechaF = datetime.strptime(fecha, '%d/%m/%Y %H:%M:%S')
        fechaF = fechaF.strftime('%Y-%m-%d %H:%M:%S')
        fechaF = datetime.strptime(fechaF, '%Y-%m-%d %H:%M:%S')
    else:
        fechaF = datetime.strptime(fecha, '%d/%m/%Y').date()
        fechaF = fechaF.strftime('%Y-%m-%d')
        fechaF = datetime.strptime(fechaF, '%Y-%m-%d').date()

    return fechaF

def montoPago(driver, engrane, mostrarColumnas, elementoAgregar, agregar, guardar):

    driver.find_element(By.XPATH, engrane).click() #Click sobre engrane
    sleep(1)
    driver.find_element(By.XPATH, mostrarColumnas).click() #Click sobre opcion "Ordenar"
    sleep(4)
    try:
        driver.find_element(By.XPATH, elementoAgregar).click() #Click sobre engrane
        sleep(1)
        driver.find_element(By.XPATH, agregar).click() #Click sobre engrane
        sleep(1)
    except:
        print('...')
        # driver.find_element(By.XPATH, '/html/body/div[20]/div[2]/div/div/div/form/table[2]/tbody/tr/td/span[3]/button').click() #Click sobre engrane
        # sleep(5)
    driver.find_element(By.XPATH, guardar).click() #Click sobre engrane
    sleep(4)

def busquedaColumna(driver, columnaBuscada):
    buscandoColumna = True
    contador = 2
    while buscandoColumna == True:
        try:
            try:
                columna = driver.find_element(By.XPATH, f"/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{str(contador)}]/div")
            except:
                columna = driver.find_element(By.XPATH, f"/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{str(contador)}]/div")
            columna = columna.text
            print('Name column : <', columna)
            sleep(5)
            if columnaBuscada in columna:
                buscandoColumna = False
                return str(contador)
            else:
                contador += 1
        except:
            print('Buscndo...')
            contador += 1

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

def validacion_cuenta_convenio_cobranza(driver, no_cuenta, no_caso):
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
            return False , 'Error Pantalla Unica', '','',''
        
        print('▬ Inician validaciones de cuenta ▬')

        sleep(5)
        labelConsultaSaldo = 'Pantalla Única de Consulta Applet de formulario:Consulta de Saldos'
        driver.find_element(By.XPATH, "//button[@aria-label='" + labelConsultaSaldo + "']").click()
        sleep(10)
        
        print('▬ Tipo Cuenta')
        try:
            tipo_cuenta = driver.find_element(By.XPATH,pantalla_unica['tipo_cuenta']['xpath'])
            tipo_cuenta = tipo_cuenta.get_attribute("value")
        except Exception:
            tipo_cuenta = driver.find_element(By.XPATH,pantalla_unica['tipo_cuenta']['xpath2'])
            tipo_cuenta = tipo_cuenta.get_attribute("value")
        # tipo_cuenta = my_copy_by_xpath(driver, pantalla_unica['tipo_cuenta']['xpath'])
        if tipo_cuenta.upper() not in ['RESIDENCIAL', 'NEGOCIO']:
            error = 'tipo cuenta ' + tipo_cuenta
            print('Error ', error)
            return False, True, '', error,''
        

        print('▬ Subtipo Cuenta')
        try:
            subtipo_cuenta = driver.find_element(By.XPATH,pantalla_unica['subtipo_cuenta']['xpath'])
            subtipo_cuenta = subtipo_cuenta.get_attribute("value")
        except Exception:
            subtipo_cuenta = driver.find_element(By.XPATH,pantalla_unica['subtipo_cuenta']['xpath2'])
            subtipo_cuenta = subtipo_cuenta.get_attribute("value")
        # subtipo_cuenta = my_copy_by_xpath(driver, pantalla_unica['subtipo_cuenta']['xpath'])
        if subtipo_cuenta.upper() not in ['NORMAL']:
            error = 'subtipo cuenta ' + subtipo_cuenta
            print('Error ', error)
            return False, True, '', error,''

        print('▬ Saldo pendiente') 
        # if open_item_selenium_wait(driver, xpath =  pantalla_unica['saldo_pendiente']['xpath']) != True: return False, True, ''

        try:
            saldo_pendiente = driver.find_element(By.XPATH,pantalla_unica['saldo_pendiente']['xpath'])
            saldo_pendiente = saldo_pendiente.get_attribute("value")
        except Exception:
            saldo_pendiente = driver.find_element(By.XPATH,pantalla_unica['saldo_pendiente']['xpath2'])
            saldo_pendiente = saldo_pendiente.get_attribute("value")
        # saldo_pendiente = my_copy_by_xpath(driver, pantalla_unica['saldo_pendiente']['xpath'])
        if ',' in saldo_pendiente:
            saldo_pendiente = saldo_pendiente.replace(',', '')
        if saldo_pendiente != None:
            try:
                saldo_pendiente = float(saldo_pendiente)
            except Exception as e:
                print(f'No se pudo convertir rl saldo a numero: {e}')
                error = 'saldo invalido'
                print('Error ', error)
                return False, True, '', error,''
        if saldo_pendiente <= 0:
            error = 'Sin Saldo Vencido'
            print('Error ', error)
            return False, True, '', error,''
        else:
            print(f'El saldo pendiente es de {saldo_pendiente} MXN')


        print('▬ Estado de la cuenta')  
        try:

            estado_cuenta = driver.find_element(By.XPATH,pantalla_unica['estado_cuenta']['xpath'])
            estado_cuenta = estado_cuenta.get_attribute("value")
        except Exception:
            estado_cuenta = driver.find_element(By.XPATH,pantalla_unica['estado_cuenta']['xpath2'])
            estado_cuenta = estado_cuenta.get_attribute("value")
        # estado_cuenta = my_copy_by_xpath(driver, pantalla_unica['estado_cuenta']['xpath'])
        if estado_cuenta.upper() not  in ['INACTIVA', 'INACTIVO', 'ACTIVO', 'ACTIVA']:
            error = 'Estado Cuenta ' + estado_cuenta
            print('Error ', error)
            return False, True, '', error,''

        sleep(4)
        print('▬ Historial de pago')

        sleep(5)

        #### ORDENAMIENTO DE TABLA 

        fecha_actual = datetime.now()
        mesAnterior = fecha_actual.replace(day=1) - timedelta(days=15)
        fechaMesAnterior = mesAnterior.replace(day=1).date()
        fechaBusqueda = datetime.strftime(fechaMesAnterior, '%d/%m/%Y')


        columnaNum = busquedaColumna(driver, 'Fecha de Pago')
        sleep(1)
        try:
            driver.find_element(By.XPATH, f"/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{columnaNum}]").click()
            colFechaPago = driver.find_element(By.XPATH, f"/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{columnaNum}]/input")
        except:
            try:
                driver.find_element(By.XPATH, f"/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{columnaNum}]").click()
                colFechaPago = driver.find_element(By.XPATH, f"/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{columnaNum}]/input")
            except:
                error = 'Error Cambio Campo Historial Pagos'
                print(error)
                return False, error, '','',''
        
        colFechaPago.send_keys(f">= '{fechaBusqueda}'")

        columnaNum = busquedaColumna(driver, 'Estado')
        sleep(1)

        try:
            driver.find_element(By.XPATH, f"/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{columnaNum}]").click()
            colFechaPago = driver.find_element(By.XPATH, f"/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{columnaNum}]/input")
        except:
            try:
                driver.find_element(By.XPATH, f"/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{columnaNum}]").click()
                colFechaPago = driver.find_element(By.XPATH, f"/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{columnaNum}]/input")
            except:
                error = 'Error Cambio Campo Historial Pagos'
                print(error)
                return False, error, '','',''
        
        colFechaPago.send_keys("Aplicado")
        colFechaPago.send_keys(Keys.RETURN)        
        colFechaPago.send_keys(Keys.RETURN)        

        print('Obteniendo la fecha del pago')
        try:
            driver.find_element(By.XPATH,  historial_pago['fechaPago']).click()
            driver.find_element(By.XPATH,  historial_pago['fechaPago']).click()
            sleep(3)
            fechaHPago= driver.find_element(By.XPATH,  historial_pago['fechaPago'] + '/input')
            fechaHPago = fechaHPago.get_attribute("value")
            print('fecha del ultimo pago: ', fechaHPago)
        except Exception:
            print('Fecha fuera de los ultimos 15 dias')
            error = 'Fecha Pago 15 dias'
            print('Error ', error)
            return False, True, '', error, ''


        try:
            sleep(5)
            driver.find_element(By.XPATH,  historial_pago['monto']).click()
            driver.find_element(By.XPATH,  historial_pago['monto']).click()
            sleep(3)
            monto= driver.find_element(By.XPATH,  historial_pago['monto'] + '/input')
            monto = monto.get_attribute("value")
            print('Monto pagado: ', monto)
        except Exception:
            print('Falla al obtener el monto del pago')
            sleep(1000)
        ajuste = 0
            # saldo_pendiente = float(saldo_pendiente)

        print('Proceso donde se realiza la obtencion del ajuste')

        monto = monto.replace('$', '')
        monto = monto.replace(',', '')

        if '150' in monto:
            rx = 'rx 150'
            if saldo_pendiente + float(monto) >= 300 and saldo_pendiente + float(monto) <= 1000:
                if '150' in monto: 
                    print('Pago correcto') 
                    ajuste = saldo_pendiente
                else: 
                    print('Pago incorrecto')
                    error = 'Pago Incorrecto'
                    print('Error ', error)
                    return False, True, '', error, rx
            else: 
                print('Rechazo por saldo Pendiente')
                error = 'Rango de Saldo'
                print('Error ', error)
                return False, True, '', error, rx
        
        elif '250' in monto:
            rx = 'rx 250'
            if saldo_pendiente + float(monto) >= 251 and saldo_pendiente + float(monto) <= 1000:
                if '250' in monto: 
                    print('Pago correcto')
                    ajuste = saldo_pendiente
                else: 
                    print('Pago incorrecto')
                    error = 'Pago Incorrecto'
                    print('Error ', error)
                    return False, True, '', error, rx
            else: 
                print('Rechazo por saldo Pendiente')
                error = 'Rango de Saldo'
                print('Error ', error)
                return False, True, '', error, rx
        
        if monto in str( saldo_pendiente * 0.3 ):
            rx = 'rx 30%'
            if saldo_pendiente + float(monto) >= 5000:
                ajuste = saldo_pendiente
            else:
                print('Rechazo por saldo Pendiente')
                error = 'Rango de Saldo'
                print('Error ', error)
                return False, True, '', error, rx

        else:
            rx = 'rx 50'
            monto = monto.replace('$', '')
            monto = monto.replace(',', '')
            if saldo_pendiente + float(monto) >= 100 and saldo_pendiente + float(monto) <= 475 or saldo_pendiente + float(monto) >= 1001 and saldo_pendiente + float(monto) <= 5000:

                if saldo_pendiente == 1500:
                    print('Rechazo por saldo Pendiente')
                    error = 'Rango de Saldo'
                    print('Error ', error)
                    return False, True, '', error, rx
                else:
                    masMenos = saldo_pendiente - float(monto)
                    if str(saldo_pendiente) in monto or (masMenos >= -1.0 and masMenos <= 1.0) : 
                        print('Pago Correcto') 
                        ajuste = saldo_pendiente
                    else: 
                        print('Pago incorrecto')
                        error = 'Pago Incorrecto'
                        print('Error ', error)
                        return False, True, '', error, rx
            else: 
                print('Rechazo por saldo Pendiente')
                error = 'Saldo Pendiente Sin Rango'
                print('Error ', error)
                return False, True, '', error, rx


        print('• Monto del ajuste: ', ajuste)
        
        print('▬ Solicitudes de ajuste')

        ##### BUSQUEDA

        sleep(2)

        try:

            driver.find_element(By.XPATH, "//button[@aria-label='Solicitudes de Ajuste Applet de lista:Consulta']").click()
            sleep(4)
        
            fechaActual = datetime.now().date()
            fechaActual = fechaActual.replace(day=1)
            fechaActual = fechaActual.strftime('%d/%m/%Y')

            columna = busquedaCol(driver, 'Motivo del ajuste', solicitudesAjustes['columnas'])
            pathMotivoAjuste = solicitudesAjustes['motivoAjuste'].replace('{contador}', columna)
            driver.find_element(By.XPATH, pathMotivoAjuste).click()
            sleep(2)
            motivoAjuste = driver.find_element(By.XPATH, pathMotivoAjuste + '/input')
            motivoAjuste.send_keys('"CARGO POR PAGO EXTEMPORANEO"')
            sleep(5)

            columna = busquedaCol(driver, 'Fecha del ajuste', solicitudesAjustes['columnas'])
            pathFechaAjuste = solicitudesAjustes['fechaAjuste'].replace('{contador}', columna)
            driver.find_element(By.XPATH, pathFechaAjuste).click()
            sleep(2)
            fechaAjuste = driver.find_element(By.XPATH, pathFechaAjuste + '/input')
            fechaAjuste.send_keys(f">= '{fechaActual}'")
            fechaAjuste.send_keys(Keys.RETURN)
            sleep(8)
            driver.find_element(By.XPATH, pathFechaAjuste).click()
            print('<<<<<<<<<<<<<<< Ajustes en Mes Actual >>>>>>>>>>>>>')
            error = 'Ajuste reciente'
            print('Error ', error)
            return False, True, '', error, rx
        
        except Exception:
            print('<<<<<<<<<<< Sin Ajustes en Mes Actual >>>>>>>>>>>>>')
            sleep(3)

        try:

            driver.find_element(By.XPATH, "//button[@aria-label='Solicitudes de Ajuste Applet de lista:Consulta']").click()
            sleep(4)
        
            fechaActual = datetime.now().date()
            fechaActual = fechaActual.replace(day=1)
            mesAnterior = fechaActual - timedelta(days=1)
            mesAnterior = mesAnterior.strftime('%d/%m/%Y')

            fecha6MesesP = datetime.now().date()
            fecha6MesesP = fecha6MesesP - timedelta(days=30*7)
            fecha6MesesP = fecha6MesesP.replace(day=1)
            fecha6MesesP = fecha6MesesP.strftime('%d/%m/%Y')

            comenntarioFecha = ">= '{}' AND <= '{}'".format(fecha6MesesP, mesAnterior)

            columna = busquedaCol(driver, 'Motivo del ajuste', solicitudesAjustes['columnas'])
            pathMotivoAjuste = solicitudesAjustes['motivoAjuste'].replace('{contador}', columna)
            driver.find_element(By.XPATH, pathMotivoAjuste).click()
            sleep(2)
            motivoAjuste = driver.find_element(By.XPATH, pathMotivoAjuste + '/input')
            motivoAjuste.send_keys('"CARGO POR PAGO EXTEMPORANEO"')
            sleep(5)

            columna = busquedaCol(driver, 'Fecha del ajuste', solicitudesAjustes['columnas'])
            pathFechaAjuste = solicitudesAjustes['fechaAjuste'].replace('{contador}', columna)
            driver.find_element(By.XPATH, pathFechaAjuste).click()
            sleep(2)
            fechaAjuste = driver.find_element(By.XPATH, pathFechaAjuste + '/input')
            fechaAjuste.send_keys(comenntarioFecha)
            fechaAjuste.send_keys(Keys.RETURN)
            fechaAjuste.send_keys(Keys.RETURN)
            sleep(8)
            driver.find_element(By.XPATH, pathFechaAjuste).click()
            print('<<<<<<<<<<<<<<< Ajustes en 6 Mes Actual >>>>>>>>>>>>>')
            error = 'Ajuste 6 Meses'
            print('Error ', error)
            return False, True, '', error, rx
        
        except Exception as e:
            print('<<<<<<<<<<< Sin Ajustes en 6 Mes >>>>>>>>>>>>>')

        sleep(5)
 
        driver.find_element(By.XPATH, "//button[@aria-label='Solicitudes de Ajuste Applet de lista:Consulta']").click()
        sleep(5)
        columna = busquedaCol(driver, 'Fecha del ajuste', solicitudesAjustes['columnas'])
        pathFechaAjuste = solicitudesAjustes['fechaAjuste'].replace('{contador}', columna)
        driver.find_element(By.XPATH, pathFechaAjuste).click()
        sleep(2)
        fechaAjuste = driver.find_element(By.XPATH, pathFechaAjuste + '/input')
        fechaAjuste.send_keys('')
        fechaAjuste.send_keys(Keys.RETURN)

        print('Validaciones correctas')
        return True, 'listo', ajuste, '', rx

    except Exception as e:
        print(f'ERROR ajustando el cargo extemporaneo. Caso NO. {no_caso}. CUENTA: {no_cuenta}')
        description_error('11','validacion_cuenta_cargo_extemporaneo',e)
        print('Error ', e)
        return False, str(e), '', 'Excepcion Validacion', ''

def extraccion_numero_pago(comentarios):
    '''
    Funcion que extrae el numero de pago de los comentarios ingresados en el ajuste

    arg:
        - comentarios: str

    out:
        - numero_pago: str
    '''
    try:
        print('Extraccion del numero de pago.')
        comentarios = comentarios.upper()
        comentarios_separados = comentarios.split('1-')
        ultimo_match = comentarios_separados[-1]
        patron = r"^[0-9]+$"
        numero_pago = '1-'
        # Verificar si la cadena coincide con la expresión regular

        for i in ultimo_match:
            if re.match(patron, i):
                numero_pago = numero_pago + i
            else:
                break
        print(f'Numero de pago {numero_pago}')
        return numero_pago
    except Exception as e:
        print('No se pudo encontrar el numero de pago.')
        description_error('CC-01','extraccion_numero_pago.',e)
        return False

def validar_ajuste_previo_cobranza(driver):
    wait = WebDriverWait(driver, 120)
    act = webdriver.ActionChains(driver)

    text_box('VALIDANDO AJUSTE CONVENIO DE COBRANZA', '▬')

def aplicacion_ajuste_convenio_cobranza(driver, caso_negocio, monto):
    '''
    Funcion que aplica el ajuste
    arg: 
        - driver
    out:
        - bool: True en caso de que se cree el ajuste, se guarde y se envie
    '''
    try:
        text_box('AJUSTE POR CONVENIO COBRANZA ', '▬')
        print(monto)
        print('Cargando...')
        open_item_selenium_wait(driver ,name = pantalla_unica['saldo_pendiente']['name'], xpath =  pantalla_unica['saldo_pendiente']['xpath'])
        it.send('{DOWN 25}')
        element = driver.find_element(By.XPATH,'//*[@id="a_5"]/div[1]')
        print('Nueva Solicitud de Ajuste')
        print('Copia el ID del ajuste')
        print('-> El monto a ajustar es: ', monto)

        open_item_selenium_wait(driver, xpath =  solicitud_ajuste['nueva_solicitud']['xpath'])
        sleep(5)
        element = driver.find_element(By.XPATH,  solicitud_ajuste['importe']['xpath'])
        element.clear()                    #Limpia lo que haya en el campo
        element.send_keys(monto)           #Introduce el numero de orden
        sleep(3)
        element = driver.find_element(By.XPATH, ajuste_CE['aplicar'])
        element.clear()                    #Limpia lo que haya en el campo
        element.send_keys('A favor')
        sleep(3)
        element = driver.find_element(By.XPATH, ajuste_CE['motivo'])
        element.clear()                    #Limpia lo que haya en el campo
        element.send_keys('CONVENIO DE COBRANZA')
        sleep(3)

        driver.find_element(By.XPATH,  solicitud_ajuste['comentario']).click()
        favor = driver.find_element(By.XPATH,  solicitud_ajuste['comentario'] + '/span/textarea')
        favor.clear()                    #Limpia lo que haya en el campo
        favor.send_keys(f'''SE APLICA AJUSTE POR CONVENIO DE COBRANZA CN {caso_negocio} ROBOT''')           #Introduce el numero de orden

        sleep(3)
        fechaAjusteGenerado= driver.find_element(By.XPATH,  '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/div/table[2]/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/span/input')
        fechaAjusteGenerado = fechaAjusteGenerado.get_attribute("value")

        sleep(3)
        print('Guardar')
        driver.find_element(By.XPATH, "//button[@aria-label='Solicitud de Ajuste Applet de formulario:Guardar']").click()
        sleep(10)
        driver.find_element(By.XPATH, "//button[@aria-label='Solicitudes de Ajuste Applet de lista:Enviar']").click()
        sleep(5)
        driver.find_element(By.XPATH, "//button[@aria-label='Enviar Ajuste Applet de formulario:Aceptar']").click()
        sleep(22)

        try:
            
            alert = Alert(driver)
            alert_txt = alert.text
            print(alert_txt)
            if 'HA SOBREPASADO EL MONTO MAXIMO DE ENVIO' in alert_txt:
                error = 'Error Siebel Sobrepasa Monto x Mes'
                alert.accept()
                driver.find_element(By.XPATH,  '/html/body/div[11]/div[2]/div/div/div/form/table/tbody/tr/td/span/span[2]/button').click()
                return False, error
        
        except Exception:

            print('Envio Correcto')
        sleep(7)


        sleep(2)
        print('Consulta saldos')
        it.send('{UP 35}')
        driver.find_element(By.XPATH, "//button[@aria-label='Pantalla Única de Consulta Applet de formulario:Consulta de Saldos']").click()
        sleep(3)
        open_item_selenium_wait(driver ,name = pantalla_unica['saldo_pendiente']['name'], xpath =  pantalla_unica['saldo_pendiente']['xpath'])
        print('AJUSTE REALIZADO CON EXITO ☺')
        
        return True, ''
        

    
    except Exception as e:
        print(f'ERROR ajustando el cargo extemporaneo. Caso NO. . CUENTA:')
        description_error('13','aplicacion_ajuste_cargo_extemporaneo',e)
        return False, 'Excepcion'

