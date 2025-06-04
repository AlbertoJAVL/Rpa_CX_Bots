from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert
from dateutil.relativedelta import relativedelta
from selenium.webdriver.common.keys import Keys
from datetime import datetime, date, timedelta
from selenium.webdriver.common.by import By
from selenium import webdriver
from time import sleep
from rutas import *
import autoit as it


import  pyautogui  as pg



def busqueda_descripcion(driver,descripcion ):
    
    detalles = '//*[@id="a_2"]/div[1]'
    sleep(2)
    open_item_selenium_wait(driver, xpath = detalles)

    i = 1
    busquedaFactura = True
    while busquedaFactura == True:
        try:
            driver.find_element(By.XPATH, f'//*[@id="{str(i)}_s_2_l_Comments"]').click()
            sleep(2)
            facturas = driver.find_element(By.XPATH, f'//*[@id="{str(i)}_Comments"]')
            facturas = facturas.get_attribute("value")

            if descripcion in facturas:
                print('♦ Ajustes Previo 6 Meses Encontrado ♦')
                print('Regresa a Pantalla Unica ')
                sleep(2)
                it.send('{UP 20}')
                open_item_selenium_wait(driver, xpath ='/html/body/div[1]/div/div[5]/div/div[8]/div[1]/div/div[3]/ul/li[1]/span/a')
                return True

            else:
                i += 1
                print('Descripcion no encontrada')
        except Exception as e:
            print(e)
            try:
                driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div/form/span/div/div[3]/div/div/div[5]/div/table/tbody/tr/td[2]/table/tbody/tr/td[3]/span').click()
                i = 1
                sleep(2)
            except Exception as e:
                return False

def open_item_selenium_wait(driver, id = None, name = None, xpath = None, clase = None, time  = 45):
    '''
    Funcion que da clics en un item según el metodo de apertura encontrado. Con espera definida

    args:
        - driver
        - Minimo uno de los siguientes:
            - id (str)
            - name
            - xpath
            - clase
    out:
        bool: True si pudo localizar y abirir el elemento indicado

    '''

    wait = WebDriverWait(driver, time)
    act = webdriver.ActionChains(driver)
    sleep(2)
    try:
        if id ==None:
            a = 2/0 #Buscar un excepcion 
        #print('Busqueda por ID')
        wait.until(EC.element_to_be_clickable((By.ID, id))).click()
        return True
    except Exception as e:
        #print(e)
        try:
            if name != None:
                    #print('Busqueda por NAME')
                    wait.until(EC.element_to_be_clickable((By.NAME, name))).click()
                    return True
            if xpath != None:
                    #print('Busqueda por XPATH')
                    wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
                    return True
            if clase != None:
                    #print('Busqueda por CLASS_NAME')
                    wait.until(EC.element_to_be_clickable((By.CLASS_NAME, clase))).click()
                    return True
        except Exception as e:
            print(f'No se pudo abrir el item especificado')
            # description_error('06','open_item_selenium_wait',e, id = id, name = name , xpath = xpath)
            return False
    

def busqueda_factura(driver, fechaInicio, fechaFin):

    try:

        historial_facturas =  '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[1]'
        factura_actual = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/a'
        

        #Historial de Facturas
        print('Entra Historial de Facturas')
        if open_item_selenium_wait(driver, xpath = historial_facturas ) != True: return False
        sleep(3)

        print('Factura Actual')
        if open_item_selenium_wait(driver, xpath = factura_actual)!= True: return False
        #Identifica periodos, esto hasta deberia ser una funcion
        sleep(5)
        it.send('{UP 50}')
        print('Scroll hacia arriba') 
        sleep(2)
        print('Recorre los items')

        #Resvia las ultimas 6 facturas
        busquedaFactura = True
        contadorBusqueda= 0
        contadorFactura = 1
        while busquedaFactura == True:
            try:
                #Valida fecha copiada
                print('Factura #',str(contadorFactura))
                
                driver.find_element(By.XPATH, f'//*[@id="{str(contadorFactura)}_s_1_l_PeriodStartDate"]').click()
                fecha_fatcura = driver.find_element(By.XPATH, f'//*[@id="{str(contadorFactura)}_PeriodStartDate"]')
                fecha_fatcura = fecha_fatcura.get_attribute("value")
                print(fecha_fatcura)

                fecha_fatcura2 = datetime.strptime(fecha_fatcura, '%d/%m/%Y').date()
                
                if fecha_fatcura2 >= fechaInicio and fecha_fatcura2 <= fechaFin:
                    status_descripcion = busqueda_descripcion(driver, "RP" )
                    if  status_descripcion == True: 
                        print('Regresa a Pantalla Unica ')
                        sleep(2)
                        it.send('{UP 20}')
                        open_item_selenium_wait(driver, xpath ='/html/body/div[1]/div/div[5]/div/div[8]/div[1]/div/div[3]/ul/li[1]/span/a')
                        return True
                    else:
                        contadorBusqueda += 1
                        contadorFactura += 1

                else:
                    print('Regresa a Pantalla Unica ')
                    sleep(2)
                    it.send('{UP 20}')
                    open_item_selenium_wait(driver, xpath ='/html/body/div[1]/div/div[5]/div/div[8]/div[1]/div/div[3]/ul/li[1]/span/a')
                    return False
                    

            except Exception as e:
                print(e)
                try:
                    driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[5]/div/table/tbody/tr/td[2]/table/tbody/tr/td[3]/span').click()
                    contadorFactura = 1
                    sleep(2)
                except Exception as e:
                    print(e)
                    print('################ FIN BUSQUEDA FACTURA ###################')
                    sleep(2)
                    return False
        
        return False

    except Exception as e:
        print(f'ERROR ajustando el cargo extemporaneo.')
        # description_error('12','busqueda_factura',e)
        error = e
        print('Error ',error)
        return False
    

def cargandoElemento(driver, elemento, atributo, valorAtributo, path = False):

    cargando = True
    contador = 0

    while cargando:

        sleep(2)
        try: 
            print('Esperando a que el elemento cargue')
            contador += 1
            if path == False: 
                driver.find_element(By.XPATH, f"//{elemento}[@{atributo}='{valorAtributo}']").click()
                return True, ''
            else:
                driver.find_element(By.XPATH, path).click()
                return True, ''
        
        except:

            try:
                print('Validando posible warning')
                alert = Alert(driver)
                alert_txt = alert.text
                print(f'♦ {alert_txt} ♦')
                if 'Cuenta en cobertura FTTH' in alert_txt: 
                    alert.accept()
                    print('Warnign aceptado, esperando elemento')
                else: return False, f'Inconsistencia Siebel: {alert_txt}'
            except:
                print('Pantalla Cargando')
                if contador == 30: return False, 'elemento no carga'

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


def busquedaCN(driver, fecha12Meses, fechaHoy, motivo, submotivo, solucion):
    # Lupa Casos de Negocio (Click)
    print(f'→ Validando CN {solucion}')
    driver.find_element(By.XPATH, "//button[@title='Casos de Negocio Applet de lista:Consulta']").click()
    sleep(10)

    # Busqueda Campo Fecha Apertura CN
    fecha_cn = obtencionColumna(driver, 'Fecha de Apertura', path_encabezados_cn)
    driver.find_element(By.XPATH, path_campo_cn.replace('{contador}', fecha_cn)).click()
    input_fecha_cn = driver.find_element(By.XPATH, f"{path_campo_cn.replace('{contador}', fecha_cn)}/input[2]")
    sleep(2)
    input_fecha_cn.send_keys(f">= '{fecha12Meses}' AND <= '{fechaHoy}'")
    print('♦ Rango de Fecha Ingresado ♦')

    # Busqueda Campo Motivo CN 
    motivo_cn = obtencionColumna(driver, 'Motivo', path_encabezados_cn)
    driver.find_element(By.XPATH, path_campo_cn.replace('{contador}', motivo_cn)).click()
    input_motivo_cn = driver.find_element(By.XPATH, f"{path_campo_cn.replace('{contador}', motivo_cn)}/input[2]")
    sleep(2)
    input_motivo_cn.send_keys(motivo)
    if motivo != '': input_motivo_cn.send_keys(Keys.RETURN)

    # Busqueda Campo Sub Motivo CN 
    subMotivo_cn = obtencionColumna(driver, 'Submotivo', path_encabezados_cn)
    driver.find_element(By.XPATH, path_campo_cn.replace('{contador}', subMotivo_cn)).click()
    input_subMotivo_cn = driver.find_element(By.XPATH, f"{path_campo_cn.replace('{contador}', subMotivo_cn)}/input[2]")
    sleep(2)
    input_subMotivo_cn.send_keys(submotivo)
    if submotivo != '': input_subMotivo_cn.send_keys(Keys.RETURN)

    # Busqueda Campo Solucion CN 
    solucion_cn = obtencionColumna(driver, 'Solución', path_encabezados_cn)
    driver.find_element(By.XPATH, path_campo_cn.replace('{contador}', solucion_cn)).click()
    input_solucion_cn = driver.find_element(By.XPATH, f"{path_campo_cn.replace('{contador}', solucion_cn)}/input[2]")
    sleep(2)
    input_solucion_cn.send_keys(solucion)
    input_solucion_cn.send_keys(Keys.RETURN)
    input_solucion_cn.send_keys(Keys.RETURN)

    print('♦ Plantilla CN Ingresada ♦')
    sleep(10)

    try: 
        # CN Reciente (click)
        driver.find_element(By.XPATH, path_resultado_cn).click()
        print('♥ CN Reciente ♥')
        return True
    except: 
        print('♥ Sin CNs Previos ♥')
        return False

def busquedaOS(driver, limiteInferior, limiteSuperior, tipoOS, cuenta):

    # Pantalla Consulta (Click)
    driver.find_element(By.XPATH, "//a[@title='Pantalla Única de Consulta']").click()
    
    # Buscando Elemento
    lupa_busqueda_cuenta = cargandoElemento(driver, 'button', 'title', 'Pantalla Única de Consulta Applet de formulario:Consulta')
    if lupa_busqueda_cuenta == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga'
    sleep(5)

    # Ingreso Cuenta
    input_busqueda_cuenta = driver.find_element(By.XPATH, "//input[@aria-label='Numero Cuenta']")
    input_busqueda_cuenta.click()
    sleep(1)
    input_busqueda_cuenta.send_keys(cuenta)
    input_busqueda_cuenta.send_keys(Keys.RETURN)
    print('♦ Cuenta Ingresada ♦')

    # Cargando Cuenta
    lupa_busqueda_cuenta = cargandoElemento(driver, 'span', 'id', 'Saldo_Vencido_Label_12')
    if lupa_busqueda_cuenta == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga'
    print('♥ Cuenta OK! ♥')
    sleep(5)

    driver.find_element(By.XPATH, "//button[@aria-label='Ordenes de Servicio Menú List']").click()
    sleep(3)

    pos_opc_nueva_consulta = obtencionColumna(driver, 'Nueva consulta              [Alt+Q]', path_opc_menu)
    driver.find_element(By.XPATH, f'{path_opc_menu.replace("{contador}", pos_opc_nueva_consulta)}/a').click()
    sleep(5)

    fecha_orden = obtencionColumna(driver, 'Fecha de la Orden', path_encabezados_ordenes_servicio)
    driver.find_element(By.XPATH, path_campo_ordenes_servicio.replace('{contador}', fecha_orden)).click()
    input_fecha_os = driver.find_element(By.XPATH, f"{path_campo_ordenes_servicio.replace('{contador}', fecha_orden)}/input[2]")
    sleep(2)
    input_fecha_os.send_keys(f">= '{limiteInferior}' AND <= '{limiteSuperior}'")

    tipo_orden = obtencionColumna(driver, 'Tipo', path_encabezados_ordenes_servicio)
    driver.find_element(By.XPATH, path_campo_ordenes_servicio.replace('{contador}', tipo_orden)).click()
    input_tipo_os = driver.find_element(By.XPATH, f"{path_campo_ordenes_servicio.replace('{contador}', tipo_orden)}/input[2]")
    sleep(2)
    input_tipo_os.send_keys(tipoOS)
    input_tipo_os.send_keys(Keys.RETURN)
    input_tipo_os.send_keys(Keys.RETURN)

    sleep(10)

    try: 
        # OS busqueda (click)
        driver.find_element(By.XPATH, path_resultado_ordenes_servicio).click()
        print('♥ OS Reciente ♥')
        return True
    except: 
        print('♥ Sin OS Previa ♥')
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
        12: "Diciembre",
    }

    nombre_mes = meses.get(int(numero_mes))
    if nombre_mes: return nombre_mes
    else: return 'Número de mes invalido'

def obtenerFechasOS(driver):

    fechas = []            
    obteniendoFechas = True
    contador = 1

    while obteniendoFechas:
        
        contador += 1
        try:
            path = path_resultado_ordenes_servicio2.replace('{contador}', str(contador))
            fechaObtenida = driver.find_element(By.XPATH, path)
            fechaObtenida = fechaObtenida.get_attribute("value")

            fechas.append(fechaObtenida[:10])

        except: return fechas

def ingresoBusquedaAjuste(driver, campoBusqueda, busqueda, pathColumnasBAjuste, pathColumnasBAjuste2, pathColumnasBAInput, pathColumnasBAInput2, pathColumnasBAInput3):
    posicion = obtencionColumna(driver, campoBusqueda, pathColumnasBAjuste, pathColumnasBAjuste2)
    if posicion == False: return False
    
    try: 
        driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicion).replace('/input[2]', '')).click()
        sleep(1)
        numeroInputajuste = driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicion))
    except: 
        try:
            driver.find_element(By.XPATH, pathColumnasBAInput2.replace('{contador}', posicion).replace('/input', '')).click()
            sleep(1)
            numeroInputajuste = driver.find_element(By.XPATH, pathColumnasBAInput2.replace('{contador}', posicion))
        except: 
            driver.find_element(By.XPATH, pathColumnasBAInput3.replace('{contador}', posicion).replace('/input', '')).click()
            sleep(1)
            numeroInputajuste = driver.find_element(By.XPATH, pathColumnasBAInput3.replace('{contador}', posicion))
    
    numeroInputajuste.send_keys(busqueda)
    numeroInputajuste.send_keys(Keys.ENTER)
    return posicion


def home(driver): driver.find_element(By.XPATH, "//a[@title='Pantalla Única de Consulta']").click()

def deteccionCertificado(driver):

    pathRGU = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[{contador}]/td[2]'

    rguBusqueda = ['TV', 'TELEFONIA', 'INTERNET']

    busqueda = True
    contador = 2
    while busqueda:
        try:
            input_rgu = driver.find_element(By.XPATH, pathRGU.replace('{contador}', str(contador)))
            rguVal =  driver.execute_script("return arguments[0].textContent;", input_rgu)

            print(f'→ RGU: {rguVal}')

            for x in rguBusqueda:
                if x in rguVal.upper():
                    rguBusqueda.remove(x)
                    break

            if len(rguBusqueda) == 0: return True, '300'
            else: contador += 1

        except:
            try: 
                driver.find_element(By.XPATH, '//*[@id="next_pager_s_4_l"]/span').click()
                contador = 2
            except Exception as e:
                if len(rguBusqueda) == 1: return True, '200'
                elif len(rguBusqueda) == 2: return True, '100'
                else: return False, 'No aplica: Sin RGU activo'

def busquedaInternet(driver, estadoOS, sinEquipo):

    pathRGUOSA = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[{contador}]/td[4]'
                
    # Cargando OS

    lupa_busqueda_cuenta, resultado = cargandoElemento(driver, 'input', 'aria-label', 'Total')
    if lupa_busqueda_cuenta == False: 
        if 'Inconsistencia' in resultado: return False, resultado
        else: return False, 'Registro pendiente'
    sleep(5)

    # Accediendo a RGU's

    try: driver.find_element(By.XPATH, "//button[@aria-label='Expandir Ítems']").click()
    except: return False, 'Registro Pendiente'

    sleep(10)

    # Ciclo de busqueda RGU Internet 80/150

    busqueda = True
    contador = 2
    while busqueda:
        try:
            input_rgu = driver.find_element(By.XPATH, pathRGUOSA.replace('{contador}', str(contador)))
            rguVal =  driver.execute_script("return arguments[0].textContent;", input_rgu)
            input_rgu = driver.find_element(By.XPATH, pathRGUOSA.replace('{contador}', str(contador)).replace('/td[4]', '/td[3]'))
            codigoAccion = driver.execute_script("return arguments[0].textContent;", input_rgu)

            print(f'→ RGU: {rguVal}')
            print(f'→ Codigo Accion: {codigoAccion}')

            if 'Internet 80 M' in rguVal or 'Internet 150 M' in rguVal:
                if 'Nuevo' in codigoAccion or 'Modificado' in codigoAccion:
                    print('RGU Activo')
                    break
                    
                else: contador += 1
            else: contador += 1

        except:
            try: 
                driver.find_element(By.XPATH, "/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[5]/div/table/tbody/tr/td[2]/table/tbody/tr/td[3]/span").click()
                contador = 2
                sleep(6)
            except: return False, 'No aplica: Sin RGU Internet'
    

    # Extraccion de Saldo a Ajustar para el caso de que la cuenta sea CON EQUIPO
    if ((estadoOS == 'Abierta') or (estadoOS == 'Completa' and sinEquipo != True)):
        lupa_busqueda_cuenta, resultado = cargandoElemento(driver, 'input', 'aria-label', 'Total')
        if lupa_busqueda_cuenta == False: 
            if 'Inconsistencia' in resultado: return False, resultado
            else: return False, 'Registro pendiente'
        sleep(5)

        totalAjustar = driver.find_element(By.XPATH, "//input[@aria-label='Total']")
        totalAjustar = totalAjustar.get_attribute("value")
        print(f'→ Monto de la factura: {totalAjustar}')
        totalAjustar = totalAjustar.replace('$', '').replace(',', '')
        return True, totalAjustar

    else: return True, ''

def obtencionMontoMesPrevio(driver, cn, sinEquipo = False):
    try:

        # print('♦  ♦')
        # print('♥  ♥')
        
        pathColumnasBCNInput2 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input'
        pathColumnasBCNInput = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input[2]'
        pathColumnasBCN = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'

        pathColumnasBOSInput = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]'
        pathColumnasBOS = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
        
        pathBOSMultiInput =    '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[3]/td[3]'
        pathBOSMultiInput2 =    '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[{contador2}]/td[{contador}]'
        
        pathOpcOS = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[2]/div/form/span/div/div[1]/div[2]/span[1]/span/ul/li[{contador}]/a'

        pathPUC = '/html/body/div[1]/div/div[5]/div/div[8]/div[1]/div/div[3]/ul/li[1]/span/a'

        print('♥ Verificando RGU ♥')
        ajusteCarga, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Casos de Negocio Applet de lista:Consulta')
        if ajusteCarga == False: 
            if 'Inconsistencia' in resultado: return False, resultado
            else: return False, 'Registro pendiente'
        sleep(15)

        # Ingreso CN
        print('→ Buscando CN')
        posicion = obtencionColumna(driver, 'Caso de Negocio', pathColumnasBCN)
        if posicion == False: return False, 'Registro pendiente'
        sleep(3)


        try: input_busqueda_cn = driver.find_element(By.XPATH, pathColumnasBCNInput.replace('{contador}', posicion))
        except: input_busqueda_cn = driver.find_element(By.XPATH, pathColumnasBCNInput2.replace('{contador}', posicion))
        input_busqueda_cn.click()
        input_busqueda_cn.send_keys(cn)
        input_busqueda_cn.send_keys(Keys.RETURN)
        print('→ CN Ingresado OK!')
        sleep(15)

        # Obtencion del estatus del CN

        posicion = obtencionColumna(driver, 'Estado', pathColumnasBCN)
        if posicion == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(3)

        try: input_estado_cn = driver.find_element(By.XPATH, pathColumnasBCNInput.replace('{contador}', posicion).replace('/input[2]', ''))
        except: input_estado_cn = driver.find_element(By.XPATH, pathColumnasBCNInput2.replace('{contador}', posicion).replace('/input', ''))
        estado_obtenido_cn = input_estado_cn.get_attribute("title")
        print(f'→ Estado CN: {estado_obtenido_cn}')
        if 'Abierto' not in estado_obtenido_cn: return False, 'No aplica: Estatus CN'

        # Obtencion de la fecha del CN

        posicion = obtencionColumna(driver, 'Fecha de Apertura', pathColumnasBCN)
        if posicion == False: return False, 'Registro pendiente'
        sleep(3)

        try: input_fecha_cn = driver.find_element(By.XPATH, pathColumnasBCNInput.replace('{contador}', posicion).replace('/input[2]', ''))
        except: input_fecha_cn = driver.find_element(By.XPATH, pathColumnasBCNInput2.replace('{contador}', posicion).replace('/input', ''))
        fecha_obtenido_cn = input_fecha_cn.get_attribute("title")
        fecha_obtenido_cn = fecha_obtenido_cn[0:10]
        print(fecha_obtenido_cn)

        # Busqueda de OS

        sleep(2)
        driver.find_element(By.XPATH, "//button[@title='Ordenes de Servicio Menú']").click()
        
        sleep(10)

        posicionOS = obtencionColumna(driver, 'Nueva consulta              [Alt+Q]', pathOpcOS)
        if posicionOS == False: return False, 'Registro pendiente'
        driver.find_element(By.XPATH, pathOpcOS.replace('{contador}', posicionOS)).click()
        
        sleep(7)

        posicionOS = obtencionColumna(driver, 'Número de Orden', pathColumnasBOS)
        if posicionOS == False: return False, 'Registro pendiente'
        sleep(3)

        # Obtencion posicion Estado de OS

        posicionEstado = obtencionColumna(driver, 'Estado', pathColumnasBOS)
        if posicionEstado == False: return False, 'Registro pendiente'
        sleep(3)

        try: 
            driver.find_element(By.XPATH, pathColumnasBOSInput.replace('{contador}', posicionEstado)).click()
            input_estado_os = driver.find_element(By.XPATH, f"{pathColumnasBOSInput.replace('{contador}', posicionEstado)}/input")
        except: 
            try: input_estado_os = driver.find_element(By.XPATH, f"{pathColumnasBOSInput.replace('{contador}', posicionEstado)}/input[2]")
            except: return False, 'Inconsistencia Siebel: Elemento diferente'

        # Ingreso de Estado
        
        # if ordenAbierta == False: input_estado_os.send_keys('Completa')
        # else: input_estado_os.send_keys('Abierta')
        # input_estado_os.send_keys(Keys.RETURN)

        input_estado_os.send_keys('"Completa" OR "Abierta"')

        posicionFecha = obtencionColumna(driver, 'Fecha de la Orden', pathColumnasBOS)
        if posicionFecha == False: return False, 'Registro pendiente'
        sleep(3)

        try: 
            driver.find_element(By.XPATH, pathColumnasBOSInput.replace('{contador}', posicionFecha)).click()
            input_fecha_os = driver.find_element(By.XPATH, f"{pathColumnasBOSInput.replace('{contador}', posicionFecha)}/input")
        except: 
            try: input_fecha_os = driver.find_element(By.XPATH, f"{pathColumnasBOSInput.replace('{contador}', posicionFecha)}/input[2]")
            except: return False, 'Inconsistencia Siebel: Elemento diferente'

        input_fecha_os.send_keys(fecha_obtenido_cn)

        posicion = obtencionColumna(driver, 'Tipo', pathColumnasBOS)
        if posicion == False: return False, 'Registro pendiente'
        sleep(3)

        try: 
            driver.find_element(By.XPATH, pathColumnasBOSInput.replace('{contador}', posicion)).click()
            input_tipo_os = driver.find_element(By.XPATH, f"{pathColumnasBOSInput.replace('{contador}', posicion)}/input")
        except: 
            try: input_tipo_os = driver.find_element(By.XPATH, f"{pathColumnasBOSInput.replace('{contador}', posicion)}/input[2]")
            except: return False, 'Inconsistencia Siebel: Elemento diferente'

        input_tipo_os.send_keys('Cambio de Servicios')
        input_tipo_os.send_keys(Keys.RETURN)
        sleep(15)

        # Acceso a la OS
        

        try: 
            driver.find_element(By.XPATH, pathBOSMultiInput).click()
            print('Multi OS')
            
            driver.find_element(By.XPATH, "//button[@title='Ordenes de Servicio Menú']").click()
        
            sleep(10)

            posicionOA = obtencionColumna(driver, 'Ordenar - Avanzado   [Ctrl+Mayús+O]', pathOpcOS)
            if posicionOA == False: return False, 'Registro pendiente'
            driver.find_element(By.XPATH, pathOpcOS.replace('{contador}', posicionOA)).click()

            ordenamientoCarga, resultado = cargandoElemento(driver, 'input', 'aria-label', 'Campo 1')
            if ordenamientoCarga == False: 
                if 'Inconsistencia' in resultado: return False, resultado
                else: return False, 'Registro pendiente'
            sleep(2)

            ordenaXEestado = driver.find_element(By.XPATH, "//input[@aria-label='Campo 1']")
            ordenaXEestado.send_keys('Fecha de la Orden')
            ordenaXEestado.send_keys(Keys.RETURN)
            sleep(1)

            # driver.find_element(By.XPATH, "//input[@rn='rdbDesc1']").click()
            # sleep(1)
            driver.find_element(By.XPATH, "//button[@aria-label='Orden de clasificación Applet de formulario:Aceptar']").click()
            sleep(10)

            # Profundizando de cada OS

            contador = 2
            while True:

                # /html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[3]
                # /html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[3]/td[3]
                try:
                    # driver.find_element(By.XPATH, pathBOSMultiInput2.replace('{contador2}', str(contador)).replace('{contador}', posicionEstado)).click()
                    numeroOS = driver.find_element(By.XPATH, pathBOSMultiInput2.replace('{contador2}', str(contador)).replace('{contador}', posicionOS))
                    numeroOS = numeroOS.get_attribute("title")

                    # Obtencion estado OS seleccionada

                    # driver.find_element(By.XPATH, pathBOSMultiInput.replace('{contador2}', str(contador)).replace('{contador}', posicionFecha)).click()
                    estadoOS = driver.find_element(By.XPATH, pathBOSMultiInput2.replace('{contador2}', str(contador)).replace('{contador}', posicionEstado))
                    estadoOS = estadoOS.get_attribute("title")

                    print(f'→ Numero OS: {numeroOS}')
                    print(f'→ Estado OS: {estadoOS}')

                    posicion = obtencionColumna(driver, 'Número de Orden', pathColumnasBOS)
                    if posicion == False: return False, 'Registro pendiente'
                    sleep(3)

                    try: 
                        
                        driver.find_element(By.XPATH, f"{pathBOSMultiInput2.replace('{contador2}', str(contador)).replace('{contador}', str(posicion))}/a").click()

                        ######################################
                        # Busqueda RGU Internet 80/150

                        resultado, totalAjustar = busquedaInternet(driver, estadoOS, sinEquipo)
                        if resultado == False: 
                            contador += 1
                            # Regresa a pantalla unica de consulta
                            driver.find_element(By.XPATH, pathPUC).click()
                            lupa_busqueda_cuenta, resultado = cargandoElemento(driver, 'span', 'id', 'Saldo_Vencido_Label_12')
                            if lupa_busqueda_cuenta == False: 
                                if 'Inconsistencia' in resultado: return False, resultado
                                else: return False, 'Registro pendiente'
                            sleep(5)

                        else: 
                            # Regresa a pantalla unica de consulta
                            driver.find_element(By.XPATH, pathPUC).click()
                            lupa_busqueda_cuenta, resultado = cargandoElemento(driver, 'span', 'id', 'Saldo_Vencido_Label_12')
                            if lupa_busqueda_cuenta == False:
                                if 'Inconsistencia' in resultado: return False, resultado
                                else: return False, 'Registro pendiente'
                            sleep(5)
                            break

                        ######################################



                    except Exception as e:  print(e); return False, 'Inconsistencia Siebel: OS Inaccesible'
                
                except Exception as e:
                    print(e)
                    try:
                        driver.find_element(By.XPATH, '//*[@id="next_pager_s_11_l]/span').click()
                        contador = 2
                    except: return False, 'No aplica: Sin RGU Internet'
            
            # Regresa a pantalla unica de consulta
            # driver.find_element(By.XPATH, pathPUC).click()
            # lupa_busqueda_cuenta = cargandoElemento(driver, 'span', 'id', 'Saldo_Vencido_Label_12')
            # if lupa_busqueda_cuenta == False: return False, 'Registro Pendiente'
            # sleep(5)

            if sinEquipo != True: return True, totalAjustar
            else:

                ######################################
                # Validacion de Certificado

                resultado, totalAjustar = deteccionCertificado(driver)
                return resultado, totalAjustar

                ######################################
            # return False, 'No aplica: Multi OS'
        except:

            try:

                # Obtencion numero OS seleccionada
                
                driver.find_element(By.XPATH, pathColumnasBOSInput.replace('{contador}', posicionEstado)).click()
                numeroOS = driver.find_element(By.XPATH, pathColumnasBOSInput.replace('{contador}', posicionOS))
                numeroOS = numeroOS.get_attribute("title")

                # Obtencion estado OS seleccionada

                driver.find_element(By.XPATH, pathColumnasBOSInput.replace('{contador}', posicionFecha)).click()
                estadoOS = driver.find_element(By.XPATH, pathColumnasBOSInput.replace('{contador}', posicionEstado))
                estadoOS = estadoOS.get_attribute("title")

                print(f'→ Numero OS: {numeroOS}')
                print(f'→ Estado OS: {estadoOS}')

                # Profundizar OS

                posicion = obtencionColumna(driver, 'Número de Orden', pathColumnasBOS)
                if posicion == False: return False, 'Registro pendiente'
                sleep(3)

                try: 
                    driver.find_element(By.XPATH, f"{pathColumnasBOSInput.replace('{contador}', posicion)}/a").click()
                except Exception as e:  print(e); return False, 'No aplica: Sin OS'
                
                ######################################
                # Busqueda RGU Internet 80/150

                resultado, totalAjustar = busquedaInternet(driver, estadoOS, sinEquipo)
                if resultado == False: return False, totalAjustar

                ######################################
                
                

                # Regresa a pantalla unica de consulta
                driver.find_element(By.XPATH, pathPUC).click()
                lupa_busqueda_cuenta, resultado = cargandoElemento(driver, 'span', 'id', 'Saldo_Vencido_Label_12')
                if lupa_busqueda_cuenta == False: 
                    if 'Inconsistencia' in resultado: return False, resultado
                    else: return False, 'Registro pendiente'
                sleep(5)

                if sinEquipo != True: return True, totalAjustar
                else:

                    ######################################
                    # Validacion de Certificado

                    resultado, totalAjustar = deteccionCertificado(driver)
                    return resultado, totalAjustar

                    ######################################


            except: return False, 'No aplica: Sin OS'

        # if ordenAbierta == False:

        #     # Ciclo de busqueda RGU Internet 80/150

        #     busqueda = True
        #     contador = 2
        #     while busqueda:
        #         try:
        #             input_rgu = driver.find_element(By.XPATH, pathRGU.replace('{contador}', str(contador)))
        #             rguVal =  driver.execute_script("return arguments[0].textContent;", input_rgu)
        #             input_rgu = driver.find_element(By.XPATH, pathRGU.replace('{contador}', str(contador)).replace('/td[2]', '/td[7]'))
        #             fechaRgu = driver.execute_script("return arguments[0].textContent;", input_rgu)
        #             fechaRgu = fechaRgu[:10]

        #             print(f'→ RGU: {rguVal}')
        #             print(f'→ Fecha RGU: {fechaRgu}')

        #             if 'Internet 80 M' in rguVal or 'Internet 150 M' in rguVal:
        #                 if fecha_obtenido_cn == fechaRgu:
        #                     print('RGU Activo')
        #                     busqueda = False
        #                 else: contador += 1
        #             else: contador += 1

        #         except:
        #             try: 
        #                 driver.find_element(By.XPATH, "//span[@title='Conjunto de registros siguiente']").click()
        #                 contador = 2
        #             except: return False, 'No aplica: Sin RGU Internet'

        #     if sinEquipo == True:

        #         # Ciclo de busqueda de rgu para aplcar certificado

        #         rguBusqueda = ['TV', 'TELEFONIA', 'INTERNET']

        #         busqueda = True
        #         contador = 2
        #         while busqueda:
        #             try:
        #                 input_rgu = driver.find_element(By.XPATH, pathRGU.replace('{contador}', str(contador)))
        #                 rguVal =  driver.execute_script("return arguments[0].textContent;", input_rgu)

        #                 print(f'→ RGU: {rguVal}')

        #                 for x in rguBusqueda:
        #                     if x in rguVal.upper():
        #                         rguBusqueda.remove(x)
        #                         break

        #                 if len(rguBusqueda) == 0: return True, '300'
        #                 else: contador += 1

        #             except:
        #                 try: 
        #                     driver.find_element(By.XPATH, '//*[@id="next_pager_s_4_l"]/span').click()
        #                     contador = 2
        #                 except Exception as e:
        #                     if len(rguBusqueda) == 1: return True, '200'
        #                     elif len(rguBusqueda) == 2: return True, '100'
        #                     else: return False, 'No aplica: Sin RGU activo'


        #     else:

            
        
        #         # driver.find_element(By.XPATH, pathColumnasBOSInput.replace('{contador}', posicionEstado)).click()
        #         # sleep(3)
        #         # driver.find_element(By.XPATH, f'//a[contains(text(), "{numeroOS}")]').click()
        #         posicion = obtencionColumna(driver, 'Número de Orden', pathColumnasBOS)
        #         if posicion == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga'
        #         sleep(3)

        #         try: 
        #             driver.find_element(By.XPATH, f"{pathColumnasBOSInput.replace('{contador}', posicion)}/a").click()
        #         except Exception as e:  print(e); return False, 'No aplica: Sin OS'
                
        #         # Cargando OS
        #         lupa_busqueda_cuenta = cargandoElemento(driver, 'input', 'aria-label', 'Total')
        #         if lupa_busqueda_cuenta == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga'
        #         sleep(5)

        #         totalAjustar = driver.find_element(By.XPATH, "//input[@aria-label='Total']")
        #         totalAjustar = totalAjustar.get_attribute("value")
        #         print(f'→ Monto para Ajustar: {totalAjustar}')

        #         driver.find_element(By.XPATH, '/html/body/div[1]/div/div[5]/div/div[8]/div[1]/div/div[3]/ul/li[1]/span/a').click()
        #         sleep(10)
                
        #         return True, totalAjustar
        
        

            


    except: 
        # sleep(10000)
        return False, 'Inconsistencia Siebel: Error FOMMP'

def reasignacionCN(driver, usuario, noCN):
    try:
        nameEncabezadosActividades = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
        inputBusquedaActividades = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input'

        lupa_busqueda_cuenta, resultado2 = cargandoElemento(driver, 'a', 'title', 'Actividades')
        if lupa_busqueda_cuenta == False: 
            if 'Inconsistencia' in resultado: return resultado2
            else: return 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(10)

        # Buscando Elemento
        lupa_busqueda_cuenta, resultado2 = cargandoElemento(driver, 'select', 'title', 'Visibilidad')
        if lupa_busqueda_cuenta == False: 
            if 'Inconsistencia' in resultado: return resultado2
            else: return 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(5)

        # Selecciona Reasignacion Manual
        resultado, resultado2 = cargandoElemento(driver, '', '', '', f'//option[contains(text(), "Reasignación manual de actividades")]')
        if resultado != True: 
            if 'Inconsistencia' in resultado: return resultado2
            else: return 'Inconsistencia Siebel: Pantalla NO Carga'
        
        # Buscando Elemento
        lupa_busqueda_cuenta, resultado2 = cargandoElemento(driver, 'button', 'title', 'Reasignar propietario Applet de lista:Consulta')
        if lupa_busqueda_cuenta == False: 
            if 'Inconsistencia' in resultado: return resultado2
            else: return 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(5)


        posicionCN = obtencionColumna(driver, 'Nº de caso de negocio', nameEncabezadosActividades)
        if posicionCN == False: return 'Inconsistencia Siebel: Pantalla NO Carga'

        driver.find_element(By.XPATH, inputBusquedaActividades.replace('{contador}', posicionCN).replace('/input', '')).click()
        input_busqueda_cn_act = driver.find_element(By.XPATH, inputBusquedaActividades.replace('{contador}', posicionCN))
        input_busqueda_cn_act.send_keys(noCN)
        input_busqueda_cn_act.send_keys(Keys.RETURN)
        sleep(15)
        try:
            posicion = obtencionColumna(driver, 'Propietario', nameEncabezadosActividades)
            if posicion == False: return 'Inconsistencia Siebel: Pantalla NO Carga'

            driver.find_element(By.XPATH, inputBusquedaActividades.replace('{contador}', posicion).replace('/input', '')).click()
            input_busqueda_prop_act = driver.find_element(By.XPATH, inputBusquedaActividades.replace('{contador}', posicion))
        except: return True
        # input_busqueda_prop_act.clear()
        sleep(5)
        if 'ftorresma' in usuario: input_busqueda_prop_act.send_keys('CVSUC2138911')
        elif 'apiliado' in usuario: input_busqueda_prop_act.send_keys('CVSUC9500410')

        posicion = obtencionColumna(driver, 'Tipo', nameEncabezadosActividades)
        if posicion == False: return 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(20)

        driver.find_element(By.XPATH, inputBusquedaActividades.replace('{contador}', posicion).replace('/input', '')).click()
        sleep(15)

        # try: driver.find_element(By.XPATH, "//a[@title='Actividades']").click()
        # except: return 'Inconsistencia Siebel: Elemento Actividades'

        return True
    except: return 'Inconsisten Siebel: Error FRCN'

# Función de apertura para generar validaciones
def inicio(driver, usuario, cuenta, noCN, tipoProceso, sinEquipo = False):

    try:
           
        pathColumnasBAjuste2 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
        pathColumnasBAjuste = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
        pathColumnasBAInput = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input[2]'
        pathColumnasBAInput2 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input'
        pathColumnasBAInput3 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input'

        print('→ Busqueda Cuenta')
        # Pantalla Casos de Negocio
        lupa_busqueda_cn, resultado = cargandoElemento(driver, 'a', 'title', 'Pantalla Única de Consulta')
        if lupa_busqueda_cn == False: 
            if 'Inconsistencia' in resultado: return resultado, '-','-','-'
            else: return 'Registro pendiente', '-','-','-'

        # Buscando Elemento
        lupa_busqueda_cn, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Pantalla Única de Consulta Applet de formulario:Consulta')
        if lupa_busqueda_cn == False: 
            if 'Inconsistencia' in resultado: return resultado, '-','-','-'
            else: return 'Registro pendiente', '-','-','-'

        lupa_busqueda_cn, resultado = cargandoElemento(driver, 'input', 'aria-label', 'Numero Cuenta')
        if lupa_busqueda_cn == False: 
            if 'Inconsistencia' in resultado: return resultado, '-','-','-'
            else: return 'Registro pendiente', '-','-','-'

        inputNCta = driver.find_element(By.XPATH, "//input[@aria-label='Numero Cuenta']")
        inputNCta.send_keys(cuenta)
        inputNCta.send_keys(Keys.RETURN)

        print('▬ Inician validaciones de cuenta ▬')

        # Cargando Cuenta
        lupa_busqueda_cn, resultado = cargandoElemento(driver, 'input', 'aria-label', 'Perfil de Pago')
        if lupa_busqueda_cn == False: 
            if 'Inconsistencia' in resultado: return resultado, '-','-','-'
            else: return 'Registro pendiente', '-','-','-'

        # Boton Consulta Saldo
        lupa_busqueda_cn, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Pantalla Única de Consulta Applet de formulario:Consulta de Saldos')
        if lupa_busqueda_cn == False: 
            if 'Inconsistencia' in resultado: return resultado, '-','-','-'
            else: return 'Registro pendiente', '-','-','-'
        sleep(15)

        # Saldo Actual
        ajustePrevio = 'No Aplica'
        btnConsultaSaldo, resultado = cargandoElemento(driver, 'input', 'aria-label', 'Saldo Total')
        if btnConsultaSaldo == False: 
            if 'Inconsistencia' in resultado: return resultado, '-','-','-'
            else: return 'Registro pendiente', '-','-','-'
        print('♥ Saldo Actualizado ♥')
        sleep(5)

        saltoTotal = driver.find_element(By.XPATH, "//input[@aria-label='Saldo Total']")
        saldoTotal = saltoTotal.get_attribute("value")
        saldoTotal = saldoTotal.replace(',', '')

        if (('-' in saldoTotal) or (float(saldoTotal) >= 0 and float(saldoTotal) <= 199)): 
            resultado, saldoTotal = obtencionMontoMesPrevio(driver, noCN, sinEquipo)
            if resultado == False: return saldoTotal, '-', '-', '-'

        elif float(saldoTotal) >=  1500: return 'No aplica: Saldo Mayor a 1500', '-', '-', '-'

        saldoTotal = float(saldoTotal.replace('$', '').replace(',', ''))
        print(sinEquipo)

        if sinEquipo == 'CC': saldoTotal = saldoTotal - 100
        if saldoTotal > 1500: return 'No aplica: Saldo Mayor a 1500', '-', '-', '-'
        
        print(f'Saldo a Ajustar: {saldoTotal}')

    
        if 'RETENCION' in tipoProceso.upper():

            pathColumnasBAjuste2 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
            pathColumnasBAjuste = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
            pathColumnasBAInput = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input[2]'
            pathColumnasBAInput2 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input'
            pathColumnasBAInput3 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input'

            fechaHoy = date.today()
            # fechaHoy = '15/03/2018'
            # fechaHoy = datetime.strptime(fechaHoy, '%d/%m/%Y').date()
            # fechaSuperior6Meses = fechaHoy.replace(day=1) - timedelta(days=1)
            fechaInferior6Meses = fechaHoy -timedelta(days=180)

            print('♥ Validando Ajuste ultimos 6 meses ♥')

            fechaSuperior6Meses2 = date.strftime(fechaHoy, '%d/%m/%Y')
            fechaInferior6Meses2 = date.strftime(fechaInferior6Meses, '%d/%m/%Y')

            rangoFechaBusqueda = f">= '{fechaInferior6Meses2}' AND <= '{fechaSuperior6Meses2}'"


            ajusteCarga, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Solicitudes de Ajuste Applet de lista:Consulta')
            if ajusteCarga == False: 
                if 'Inconsistencia' in resultado: return resultado, '-','-','-'
                else: return 'Registro pendiente', '-','-','-'
            sleep(10)

            resultado = ingresoBusquedaAjuste(driver, 'Estado', 'Aplicado', pathColumnasBAjuste, pathColumnasBAjuste2, pathColumnasBAInput, pathColumnasBAInput2, pathColumnasBAInput3)
            if resultado == False: return 'Inconsistencia Siebel: Pantalla NO Carga', '-', '-', ajustePrevio
            
            resultado = ingresoBusquedaAjuste(driver, 'Motivo del ajuste', 'RETENCION', pathColumnasBAjuste, pathColumnasBAjuste2, pathColumnasBAInput, pathColumnasBAInput2, pathColumnasBAInput3)
            if resultado == False: return 'Inconsistencia Siebel: Pantalla NO Carga', '-', '-', ajustePrevio
            
            resultado = ingresoBusquedaAjuste(driver, 'Fecha del ajuste', rangoFechaBusqueda, pathColumnasBAjuste, pathColumnasBAjuste2, pathColumnasBAInput, pathColumnasBAInput2, pathColumnasBAInput3)
            if resultado == False: return 'Inconsistencia Siebel: Pantalla NO Carga', '-', '-', ajustePrevio
            

            print('♥ Buscando Ajuste ♥')
            sleep(10)

            try: 
                driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', resultado).replace('/input[2]', '')).click()
                ajustePrevio = 'Aplica'
                print('Ajuste Previo Validacion 1 Encontrado')
                return 'No aplica: Ajuste Previo', '-', '-', ajustePrevio
            except: 
                try:
                    driver.find_element(By.XPATH, pathColumnasBAInput2.replace('{contador}', resultado).replace('/input', '')).click()
                    ajustePrevio = 'Aplica'
                    print('Ajuste Previo Validacion 1 Encontrado')
                    return 'No aplica: Ajuste Previo', '-', '-', ajustePrevio

                except: 
                    try: 
                        driver.find_element(By.XPATH, pathColumnasBAInput3.replace('{contador}', resultado).replace('/input', '')).click()
                        ajustePrevio = 'Aplica'
                        print('Ajuste Previo Validacion 1 Encontrado')
                        return 'No aplica: Ajuste Previo', '-', '-', ajustePrevio
                    except: pass

            if ajustePrevio == 'No Aplica' and sinEquipo == False: 
                resultadoBusquedaAjuPrev = busqueda_factura(driver, fechaInferior6Meses, fechaHoy)
                print('Valores despues de la busqueda de facturas')
                print(resultadoBusquedaAjuPrev)
                if resultadoBusquedaAjuPrev == True: 
                    ajustePrevio = 'Aplica'
                    print('Ajuste Previo Validacion 2 Encontrado')
                    return 'No aplica: Ajuste Previo', '-', '-', ajustePrevio

                else: print('Sin Ajustes Previos')
       

        # Generacion Ajuste
        print('Generando Ajuste')
        ajusteCarga, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Solicitudes de Ajuste Applet de lista:Nuevo')
        if ajusteCarga == False: 
            if 'Inconsistencia' in resultado: return resultado, '-','-','-'
            else: return 'Registro pendiente', '-','-','-'
        
        ajusteCarga, resultado = cargandoElemento(driver, 'input', 'aria-label', 'Importe del ajuste')
        if ajusteCarga == False: 
            if 'Inconsistencia' in resultado: return resultado, '-','-','-'
            else: return 'Registro pendiente', '-','-','-'
        print('♥ Generando Ajuste ♥')
        
        # Importe Ajuste
        input_importe = driver.find_element(By.XPATH, "//input[@aria-label='Importe del ajuste']")
        input_importe.click()
        sleep(1)
        input_importe.clear()
        input_importe.send_keys(str(saldoTotal))
        print('♦ Importe Ingresado ♦')
        # sleep(1000)

        input_aplicar = driver.find_element(By.XPATH, "//input[@aria-label='Aplicar']")
        input_aplicar.click()
        sleep(1)
        input_aplicar.send_keys('A favor')
        input_aplicar.send_keys(Keys.RETURN)
        print('♦ Aplicacion Ingresada ♦')

        input_motivo_ajuste = driver.find_element(By.XPATH, "//input[@aria-label='Motivo del ajuste']")
        input_motivo_ajuste.click()
        sleep(1)
        input_motivo_ajuste.send_keys('RETENCION')
        input_motivo_ajuste.send_keys(Keys.RETURN)
        print('♦ Motivo Ajuste Ingresado ♦')

        input_comentario = driver.find_element(By.XPATH, "//textarea[@aria-label='Comentarios']")
        input_comentario.click()
        sleep(1)
        input_comentario.send_keys(f'RETENCION 0 AJUSTE MAS MIGRACION CN {noCN}')
        print('♦ Comentario Ingresado ♦')

        numeroAjuste = driver.find_element(By.XPATH, "//input[@aria-label='Número de Ajuste']")
        numeroAjuste = numeroAjuste.get_attribute("value")
        print(f'Ajuste Generado: {numeroAjuste}')

        driver.find_element(By.XPATH, "//button[@aria-label='Solicitud de Ajuste Applet de formulario:Guardar']").click()
        print('♦ Ajuste Guardado ♦')
        sleep(5)
        
        # Envio Ajuste
        print('♥ Enviando Ajuste ♥')
        ajusteCarga, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Solicitudes de Ajuste Applet de lista:Consulta')
        if ajusteCarga == False: 
            if 'Inconsistencia' in resultado: return resultado, '-','-','-'
            else: return 'Inconsistencia Siebel: Pantalla NO Carga', '-','-','-'

        # sleep(10000)
        posicion = obtencionColumna(driver, 'Número de Ajuste', pathColumnasBAjuste, pathColumnasBAjuste2)
        if posicion == False: return 'Inconsistencia Siebel: Pantalla NO Carga', numeroAjuste, saldoTotal, ajustePrevio
        
        try: numeroInputajuste = driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicion))
        except: 
            try: numeroInputajuste = driver.find_element(By.XPATH, pathColumnasBAInput2.replace('{contador}', posicion))
            except: numeroInputajuste = driver.find_element(By.XPATH, pathColumnasBAInput3.replace('{contador}', posicion))
        
        numeroInputajuste.click()
        numeroInputajuste.send_keys(numeroAjuste)
        numeroInputajuste.send_keys(Keys.ENTER)
        print('♥ Buscando Ajuste ♥')
        sleep(5)

        # driver.find_element(By.XPATH, "//button[@aria-label='Solicitudes de Ajuste Applet de lista:Enviar']").click()
        
        ajusteCarga, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Solicitudes de Ajuste Applet de lista:Enviar')
        if ajusteCarga == False: 
            if 'Inconsistencia' in resultado: return resultado, numeroAjuste, saldoTotal, ajustePrevio
            else: return 'Inconsistencia Siebel: Pantalla NO Carga', numeroAjuste, saldoTotal, ajustePrevio

        ajusteCarga, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Enviar Ajuste Applet de formulario:Aceptar')
        if ajusteCarga == False: 
            if 'Inconsistencia' in resultado: return resultado, numeroAjuste, saldoTotal, ajustePrevio
            else: return 'Inconsistencia Siebel: Pantalla NO Carga', numeroAjuste, saldoTotal, ajustePrevio

        print('♦ Ajuste Enviado ♦')
        sleep(10)
        print('inicio reasignacion')
        resultado = reasignacionCN(driver, usuario, noCN)
        if resultado != True: return resultado, '', '-', '-' 
        print('Fin Reasignacion')

        return 'Completado', numeroAjuste, saldoTotal, ajustePrevio

    except Exception as e: print(e); return 'Error FI', '-', '-', '-'

def cierreActividad(driver, cuenta, cn, tipoProceso):

    try:
        # print('♦  ♦')
        # print('♥  ♥')
        
        pathColumnasBCN = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
        pathColumnasBCNInput = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input[2]'
        pathColumnasBCNInput2 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input'

        pathColumnasBActividad = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
        pathColumnasBAInput = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input[2]'
        # pathColumnasBAInput2 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input'

        pathOpcEstadosAc = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[5]/span'
        pathOpcEstadosAc2 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[4]/span'
        print('→ Buscando Cuenta')

        lupa_busqueda_cuenta, res = cargandoElemento(driver, 'a', 'title', 'Pantalla Única de Consulta')
        if lupa_busqueda_cuenta == False: return 'Inconsistencia Siebel: Pantalla NO Carga', '-', '-', '-'
        
        # Buscando Elemento
        lupa_busqueda_cuenta, res = cargandoElemento(driver, 'button', 'title', 'Pantalla Única de Consulta Applet de formulario:Consulta')
        if lupa_busqueda_cuenta == False: return 'Inconsistencia Siebel: Pantalla NO Carga', '-', '-', '-'
        sleep(5)

        # Ingreso Cuenta
        input_busqueda_cuenta = driver.find_element(By.XPATH, "//input[@aria-label='Numero Cuenta']")
        input_busqueda_cuenta.click()
        sleep(1)
        input_busqueda_cuenta.send_keys(cuenta)
        input_busqueda_cuenta.send_keys(Keys.RETURN)
        print('♦ Cuenta Ingresada ♦')

        # Cargando Cuenta
        lupa_busqueda_cuenta, res = cargandoElemento(driver, 'span', 'id', 'Saldo_Vencido_Label_12')
        if lupa_busqueda_cuenta == False: return 'Inconsistencia Siebel: Pantalla NO Carga', '-', '-', '-'
        print('♥ Cuenta OK! ♥')
        sleep(5)
        print('♥ Cerrando Actividad y CN ♥')
        ajusteCarga, res = cargandoElemento(driver, 'button', 'aria-label', 'Casos de Negocio Applet de lista:Consulta')
        if ajusteCarga == False: return 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(15)

        try:
            alert = Alert(driver)
            alert_txt = alert.text
            print(f'♦ {alert_txt} ♦')
            alert.accept()
            return f'Inconsistencia Siebel: {alert_txt}'
        except: pass

        # Ingreso CN
        print('→ Cierre Actividad')
        posicion = obtencionColumna(driver, 'Caso de Negocio', pathColumnasBCN)
        if posicion == False: return 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(3)


        try: input_busqueda_cn = driver.find_element(By.XPATH, pathColumnasBCNInput.replace('{contador}', posicion))
        except: input_busqueda_cn = driver.find_element(By.XPATH, pathColumnasBCNInput2.replace('{contador}', posicion))
        input_busqueda_cn.click()
        input_busqueda_cn.send_keys(cn)
        input_busqueda_cn.send_keys(Keys.RETURN)
        print('→ CN Ingresado OK!')
        sleep(15)

        posicionEstado = obtencionColumna(driver, 'Estado', pathColumnasBCN)
        if posicionEstado == False: return 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(3)

        if 'CONVENIO' in tipoProceso.upper():
            posicion = obtencionColumna(driver, 'Comentarios', pathColumnasBCN)
            if posicion == False: return 'Inconsistencia Siebel: Pantalla NO Carga'
            sleep(3)

            try: input_comentarios_cn = driver.find_element(By.XPATH, pathColumnasBCNInput.replace('{contador}', posicion).replace('/input[2]', ''))
            except: input_comentarios_cn = driver.find_element(By.XPATH, pathColumnasBCNInput2.replace('{contador}', posicion).replace('/input', ''))
            comentario_obtenido_cn = input_comentarios_cn.get_attribute("title")
            print(comentario_obtenido_cn)


        try: driver.find_element(By.XPATH, pathColumnasBCNInput.replace('{contador}', posicionEstado).replace('/input[2]', '')).click()
        except: driver.find_element(By.XPATH, pathColumnasBCNInput2.replace('{contador}', posicionEstado).replace('/input', '')).click()
        sleep(1)

        resultado,res = cargandoElemento(driver, '', '', '', f'//a[contains(text(), "{cn}")]')
        if resultado != True: return 'Inconsistencia Siebel: Pantalla NO Carga'
        print('♦ CN Abierto ♦')

        # Actividades
        resultado,res = cargandoElemento(driver, '', '', '', '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/ul/li[2]/a')
        if resultado != True: return 'Inconsistencia Siebel: Pantalla NO Carga'
        print('♦ Actividades Abiertas ♦')
        sleep(10)

        # Comentarios
        posicion = obtencionColumna(driver, 'Comentarios', pathColumnasBActividad)
        if posicion == False: return 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(3)

        driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicion).replace('/input[2]', '')).click()
        sleep(1)
        input_busqueda_comentarios = driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicion).replace('/input[2]', '/textarea'))
        try: input_busqueda_comentarios.click()
        except: return 'Inconsistencia Siebel: Asignacion'

        if 'CONVENIO' in tipoProceso.upper(): input_busqueda_comentarios.send_keys(comentario_obtenido_cn)
        else: input_busqueda_comentarios.send_keys(f'RETENCION 0 AJUSTE MAS MIGRACION CN {cn}')
        print('♥ Comentario Actividad Ingresado ♥')

        # Motivo Cierre
        posicion = obtencionColumna(driver, 'Motivo del cierre', pathColumnasBActividad)
        if posicion == False: return 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(3)

        try:
            driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicion).replace('/input[2]', '')).click()
            input_busqueda_mcierre = driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicion))
            input_busqueda_mcierre.send_keys('SE APLICA AJUSTE')
            input_busqueda_mcierre.send_keys(Keys.RETURN)
            print('♥ Motivo Cierre Actividad Ingresado ♥')
        except Exception:
            try:
                # driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicion).replace('/input[2]', '')).click()
                input_busqueda_mcierre = driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicion).replace('/input[2]', '/input'))
                input_busqueda_mcierre.send_keys('SE APLICA AJUSTE')
                input_busqueda_mcierre.send_keys(Keys.RETURN)
                print('♥ Motivo Cierre Actividad Ingresado ♥')
            except Exception as e: print(f'{str(e)}');  return 'Inconsistencia Siebel: Cambio Elementos'

        

        # Estado Actividad
        posicion = obtencionColumna(driver, 'Estado', pathColumnasBActividad)
        if posicion == False: return 'Inconsistencia Siebel: Pantalla NO Carga'
        
        driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicion).replace('/input[2]', '')).click()
        sleep(2)

        try: driver.find_element(By.XPATH, pathOpcEstadosAc).click()
        except:
            try: driver.find_element(By.XPATH, pathOpcEstadosAc2).click()
            except Exception as e: print(str(e)); return 'Inconsistencia Siebel: Cambio Elementos'

        resultado,res = cargandoElemento(driver, '', '', '', '//div[contains(text(), "CERRADA")]')
        if resultado != True: return 'Inconsistencia Siebel: Pantalla NO Carga'

        print('♥ Motivo Cierre Actividad Ingresado ♥')
        driver.find_element(By.XPATH, "//button[@aria-label='Caso de negocio Applet de formulario:Guardar']").click()
        sleep(15)

        try:
            alert = Alert(driver)
            alert_txt = alert.text
            print(type(alert_txt))
            print(f'♦ {alert_txt} ♦')
            alert.accept()
            if 'WriteRecord no se permite' in str(alert_txt): pass
            else: return f'Inconsistencia Siebel: {alert_txt}'

        except: pass
        sleep(5)

        print('♦ Cuenta terminada ♦')

        
        return 'Completado'
        
    except Exception as e: 
        print(f'ERROR EN FUNCION INICIO. ERROR: {e}')
        return 'FCierreOS'
    
