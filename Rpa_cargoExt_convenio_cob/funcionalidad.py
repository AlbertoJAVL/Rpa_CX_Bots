from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert
from dateutil.relativedelta import relativedelta
from selenium.webdriver.common.keys import Keys
from datetime import datetime, date, timedelta
from selenium.webdriver.common.by import By
from datetime import datetime, date, timedelta

from time import sleep
import os


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

def cargandoElemento(driver, elemento, atributo, valorAtributo, path = False):

    cargando = True
    contador = 0

    while cargando:

        sleep(1)
        try: 
            print('Validando posible warning')
            contador += 1
            sleep(1)
            alert = Alert(driver)
            alert_txt = alert.text
            print(f'→→ {alert_txt}')
            if 'Cuenta con cobertura FTTH' in alert_txt:
                alert.accept()
                print('Warningn validado')
                # return True, ''
            else: return False, f'Inconsistencia Siebel: {alert_txt}'
            
            
        
        except:
            try:
                print('Esperando a que el elemento cargue')
                if path == False: 
                    driver.find_element(By.XPATH, f"//{elemento}[@{atributo}='{valorAtributo}']").click()
                    return True, ''
                else: 
                    driver.find_element(By.XPATH, path).click()
                    return True, ''
            except:
                print('Pantalla Cargando')
                if contador == 60: return False, ''

def cargandoResultado(driver, path):
    
    contador = 0
    while True:
        try:
            
            contador += 1
            sleep(1)
            driver.find_element(By.XPATH, path).click()
            return True
        
        except:
            if contador == 20: return False

# def extraer_texto_de_pdf_con_imagenes(ruta_pdf):
#     try:
#         print('Iniciando LEctura PDF')
#         print(ruta_pdf)
#         # Instancia de tesseract
#         # pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
#         pytesseract.pytesseract.tesseract_cmd = r'C:\Users\sistemas\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
#         print(1)
#         # Instancia de Poppler
#         poppler_path = r'C:\Program Files\poppler-24.07.0\Library\bin'
#         print(2)

#         # Convertir PDF a imagenes
#         imagenes = convert_from_path(ruta_pdf, poppler_path=poppler_path)

#         # Inicio de extraccion de texto

#         texto_completo = ""

#         for i, imagen in enumerate(imagenes):
            
#             temp_image = f"temp_image_{i}.png"
#             imagen.save(temp_image, 'PNG')
#             texto = pytesseract.image_to_string(Image.open(temp_image))
#             texto_completo += f"Página {i+1}:\n{texto}\n\n"
#             os.remove(temp_image)

#             if "CARGO POR PAGO TARDIO" in texto_completo.upper():
#                 print('Cargo detectado')

#                 texto = texto_completo.upper().split("CARGO POR PAGO TARDIO ", 1)
#                 print(texto)
#                 if len(texto) > 1:
#                     importeAjuste = texto[1]
#                     importeAjuste = importeAjuste.splitlines()[0].strip()
#                     if '$' in importeAjuste:
#                         posicionImporte = importeAjuste.find('$')
#                         importeAjuste = importeAjuste[posicionImporte:]
#                     if '.' in importeAjuste:
#                         posicionImporte2 = importeAjuste.find('.')
#                         importeAjuste = importeAjuste[:posicionImporte2+2]
#                     print(f'Monto del pago tardío: {importeAjuste}')
#                     # sleep(10000)
#                     return True, importeAjuste
#                 else:
#                     print('Monto de Pago Tardío no detectado')
#                     return False, 'Monto Pago No Detectado'

#             else: print('NO')
#         return False, 'Sin Pago Tardio'

#     except Exception as e: print(e); sleep(10000)
    

def validacionTipoCN(driver, noCN):
    try:

        pathEncabezadosPCN = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div' 
        pathRegistrosPCN = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input[2]'

        print('→ INICIO DE VALIDACIONES DE CN')
        print('→ Buscando CN')

        # Pantalla Casos de Negocio
        lupa_busqueda_cn, resultado = cargandoElemento(driver, 'a', 'title', 'Casos de negocio')
        if lupa_busqueda_cn == False: return False, 'Pendiente'

        # Buscando Elemento
        lupa_busqueda_cn, resultado = cargandoElemento(driver, 'button', 'title', 'Casos de negocio Applet de lista:Consulta')
        if lupa_busqueda_cn == False: return False, 'Pendiente'
        sleep(5)

        print('→ Buscando CN')

        posicionStaCN = obtencionColumna(driver, 'Estado', pathEncabezadosPCN)
        if posicionStaCN == False: return False, 'Pendiente'

        posicionCN = obtencionColumna(driver, 'Caso de negocio', pathEncabezadosPCN)
        if posicionCN == False: return False, 'Pendiente'
        sleep(3)

        driver.find_element(By.XPATH, pathRegistrosPCN.replace('{contador}', posicionCN).replace('/input[2]', '')).click()
        try: inputCN = driver.find_element(By.XPATH, pathRegistrosPCN.replace('{contador}', posicionCN))
        except: inputCN = driver.find_element(By.XPATH, pathRegistrosPCN.replace('{contador}', posicionCN).replace('/input[2]', '/input'))

        inputCN.send_keys(noCN)
        inputCN.send_keys(Keys.RETURN)
        sleep(5)

        try:

            valorEstatusCN = driver.find_element(By.XPATH, pathRegistrosPCN.replace('{contador}', posicionStaCN).replace('/input[2]', ''))
            valorEstatusCN = valorEstatusCN.get_attribute('title')
            print(f'→ Estatus Inicial CN: {valorEstatusCN}')

            if 'Cerrado' in valorEstatusCN: return True, 'Cerrado'
            elif 'Cancelado' in valorEstatusCN: return True, 'Cancelado'
            else: return True, ''

        except: return True, 'Pendiente'



    except Exception as e: 
        return True, f'Error: {str(e)}'
    
def validacionCargoExt(driver, noCta, noCN):

    try: 

        pathEncabezadoBAjustes = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
        pathRegistroBAjustes = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input[2]'

        pathEncabezadoBCN = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
        pathRegistroBCN = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input[2]'

        # /html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[3]

        print('→ INICIO VALIDACIONES CARGO EXTEMPORANEO')
        print('→ Busqueda Cuenta')
        # Pantalla Casos de Negocio
        lupa_busqueda_cn, resultado = cargandoElemento(driver, 'a', 'title', 'Pantalla Única de Consulta')
        if lupa_busqueda_cn == False: return False, 'Pendiente'

        # Buscando Elemento
        lupa_busqueda_cn, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Pantalla Única de Consulta Applet de formulario:Consulta')
        if lupa_busqueda_cn == False: return False, 'Pendiente'

        lupa_busqueda_cn, resultado = cargandoElemento(driver, 'input', 'aria-label', 'Numero Cuenta')
        if lupa_busqueda_cn == False: return False, 'Pendiente'

        inputNCta = driver.find_element(By.XPATH, "//input[@aria-label='Numero Cuenta']")
        inputNCta.send_keys(noCta)
        inputNCta.send_keys(Keys.RETURN)

        print('▬ Inician validaciones de cuenta ▬')

        lupa_busqueda_cn, resultado = cargandoElemento(driver, 'input', 'aria-label', 'Perfil de Pago')
        if lupa_busqueda_cn == False: return False, 'Pendiente'

        # Tipo Cuenta
        tipo_cuenta = driver.find_element(By.XPATH, "//input[@aria-label='Tipo']")
        tipo_cuenta = tipo_cuenta.get_attribute("value")

        print(f'Tipo Cuenta: {tipo_cuenta}')
        if tipo_cuenta.upper() not in ['RESIDENCIAL', 'NEGOCIO']:
             error = f'No aplica: Tipo Cuenta {tipo_cuenta}'
             print(error)
             return False, error
        
        # Subtipo Cuenta
        subtipo_cuenta = driver.find_element(By.XPATH, "//input[@aria-label='SubTipo']")
        subtipo_cuenta = subtipo_cuenta.get_attribute("value")
        
        print(f'Subtipo Cuenta: {subtipo_cuenta}')
        if subtipo_cuenta.upper() not in ['NORMAL']:
            error = f'No aplica: Subtipo Cuenta {subtipo_cuenta}'
            print(error)
            return False, error
        
        # Validacion Ajustes Previos
        # Esc. Mes Actual
        print('→ Validando Ajuste Reciente')

        lupa_busqueda_a, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Solicitudes de Ajuste Applet de lista:Consulta')
        if lupa_busqueda_a == False: return False, 'Pendiente'
        sleep(5)

        fechaActual = datetime.now().date()
        fechaActual = fechaActual.replace(day=1)
        fechaActual = fechaActual.strftime('%d/%m/%Y')
        comenntarioFecha = ">= '" + fechaActual + "'"

        posicionFA = obtencionColumna(driver, 'Fecha del ajuste', pathEncabezadoBAjustes)
        if posicionFA == False: return False, 'Pendiente'

        posicionEA = obtencionColumna(driver, 'Estado', pathEncabezadoBAjustes)
        if posicionEA == False: return False, 'Pendiente'

        posicionMA = obtencionColumna(driver, 'Motivo del ajuste', pathEncabezadoBAjustes)
        if posicionMA == False: return False, 'Pendiente'

        # Fecha Ajuste
        driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionFA).replace('/input[2]', '')).click()
        
        try: inputFA = driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionFA))
        except: inputFA = driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionFA).replace('/input[2]', '/input'))

        inputFA.send_keys(comenntarioFecha)

        # Estado Ajuste
        driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionEA).replace('/input[2]', '')).click()
        
        try: inputEA = driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionEA))
        except: inputEA = driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionEA).replace('/input[2]', '/input'))

        inputEA.send_keys('Aplicado')
        inputEA.send_keys(Keys.RETURN)

        # Motivo Ajuste
        driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionMA).replace('/input[2]', '')).click()
        
        try: inputMA = driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionMA))
        except: inputMA = driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionMA).replace('/input[2]', '/input'))

        inputMA.send_keys('CARGO POR PAGO EXTEMPORANEO')
        inputMA.send_keys(Keys.RETURN)
        inputMA.send_keys(Keys.RETURN)

        resultadoCarga = cargandoResultado(driver, pathRegistroBAjustes.replace('{contador}', posicionEA).replace('/input[2]', ''))
        if resultadoCarga == True: 
            error = 'No aplica: Ajuste Mes'
            print(error)
            return False, error
        
        print('→ Sin Ajustes en el mes corriente')

        # Esc. Ultimos 3 meses
        print('→ Validando Ajustes en los ultimos 3 meses')
        lupa_busqueda_a, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Solicitudes de Ajuste Applet de lista:Consulta')
        if lupa_busqueda_a == False: return False, 'Pendiente'
        sleep(5)

        fechaActual = date.today()
        fechaActual = fechaActual.replace(day=1)
        mesAnterior = fechaActual - timedelta(days=1)
        mesAnterior = mesAnterior.strftime('%d/%m/%Y')

        fecha6MesesP = date.today()
        fecha6MesesP = fecha6MesesP - timedelta(days=30*3)
        fecha6MesesP = fecha6MesesP.replace(day=1)
        fecha6MesesP = fecha6MesesP.strftime('%d/%m/%Y')

        comenntarioFecha = ">= '{}' AND <= '{}'".format(fecha6MesesP, mesAnterior)

        # Fecha Ajuste
        driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionFA).replace('/input[2]', '')).click()
        
        try: inputFA = driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionFA))
        except: inputFA = driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionFA).replace('/input[2]', '/input'))

        inputFA.send_keys(comenntarioFecha)

        # Estado Ajuste
        driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionEA).replace('/input[2]', '')).click()
        
        try: inputEA = driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionEA))
        except: inputEA = driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionEA).replace('/input[2]', '/input'))

        inputEA.send_keys('Aplicado')
        inputEA.send_keys(Keys.RETURN)

        # Motivo Ajuste
        driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionMA).replace('/input[2]', '')).click()
        
        try: inputMA = driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionMA))
        except: inputMA = driver.find_element(By.XPATH, pathRegistroBAjustes.replace('{contador}', posicionMA).replace('/input[2]', '/input'))

        inputMA.send_keys('CARGO POR PAGO EXTEMPORANEO')
        inputMA.send_keys(Keys.RETURN)
        inputMA.send_keys(Keys.RETURN)

        resultadoCarga = cargandoResultado(driver, pathRegistroBAjustes.replace('{contador}', posicionEA).replace('/input[2]', ''))
        if resultadoCarga == True: 
            error = 'No aplica: Ajuste Reciente'
            print(error)
            return False, error

        print('→ Sin Ajustes en los ultimos 3 meses')

        # Validacion CN Sucursal

        print('→ Validando CN Sucursarl')

        lupa_busqueda_cn, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Casos de Negocio Applet de lista:Consulta')
        if lupa_busqueda_cn == False: return False, 'Pendiente'
        sleep(5)

        posicionMCCN = obtencionColumna(driver, 'Motivo Cliente', pathEncabezadoBCN)
        if posicionMCCN == False: return False, 'Pendiente'

        # Motivo Ajuste
        driver.find_element(By.XPATH, pathRegistroBCN.replace('{contador}', posicionMCCN).replace('/input[2]', '')).click()
        
        try: inputMA = driver.find_element(By.XPATH, pathRegistroBCN.replace('{contador}', posicionMCCN))
        except: inputMA = driver.find_element(By.XPATH, pathRegistroBCN.replace('{contador}', posicionMCCN).replace('/input[2]', '/input'))

        inputMA.send_keys('CARGO EXTEMPORANEO')
        inputMA.send_keys(Keys.RETURN)
        inputMA.send_keys(Keys.RETURN)

        resultadoCarga = cargandoResultado(driver, pathRegistroBCN.replace('{contador}', posicionMCCN).replace('/input[2]', ''))
        if resultadoCarga == False: 
            error = 'No aplica: Sin CN Sucursal'
            print(error)
            return False, error
        
        print('Validaciones Completa c:')
        return True, ''

    except Exception as e: print(e); sleep(10000)

def deteccion_cargo(driver):

    pathItemFactura = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[{contador}]/td[3]'
    pathValItemFact = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[{contador}]/td[4]'

    pathDesplazaItems = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div/form/span/div/div[3]/div/div/div[5]/div/table/tbody/tr/td[2]/table/tbody/tr/td[3]'
    
    contador = 2
    while True:

        try:

            sleep(0.75)
            driver.find_element(By.XPATH, pathItemFactura.replace('{contador}', str(contador))).click()
            itemFacturacion = driver.find_element(By.XPATH, pathItemFactura.replace('{contador}', str(contador)))
            itemFacturacion = itemFacturacion.get_attribute('title')
            print(f'→ Item Factura: {itemFacturacion}')

            if 'Cargo por pago tardío:' in itemFacturacion:
                valItem = driver.find_element(By.XPATH, pathValItemFact.replace('{contador}', str(contador)))
                valItem = valItem.get_attribute('title')

                print(f'→ Valor detectado: {valItem}')
                return valItem
            
            else: contador += 1
                

        except: 
            try: 
                driver.find_element(By.XPATH, pathDesplazaItems).click()
                contador = 2
            except: return False

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

    pathFacturaActual = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/a'
    pathEncabezadoFacturas = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
    pathFacturas = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[{contador2}]/td[{contador}]'
    pathSigFactura = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[5]/div/table/tbody/tr/td[2]/table/tbody/tr/td[3]/span'
    


    # pathCuadroPDF = '/html/body/div[21]/div[2]'
    # pathEmbedPDF = '/html/body/div[21]/div[2]/embed'
    # pathCierrePDF = '/html/body/div[21]/div[1]/button'

    
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

        print('▬ BUSQUEDA DE FACTURA ▬')

        # Profundizar en las facturas
        profundizar_factura, resultado = cargandoElemento(driver, '', '', '', pathFacturaActual)
        if profundizar_factura == False: return False, '-', 'Pendiente', '-'

        # Validacion de historico de facturas
        facturas_activas, resultado = cargandoElemento(driver, '', '', '', '//div[contains(text(), "Historial de facturas")]')
        if facturas_activas == False: return False, '-', 'Pendiente', '-'

        # Obtencion de posicion de la columna fecha
        posFecha = obtencionColumna(driver, 'Fecha inicial del periodo de facturación', pathEncabezadoFacturas)
        if posFecha == False: return False, '-', 'Pendiente', '-'

        fechaHoy = datetime.today()
        rango3Meses = fechaHoy - timedelta(days=91)

        print('Validando Facturas')
        
        # Busqueda de elemento
        contador = 2
        while True: 
            try:
                
                
                # Obtencion Fecha de Factura
                try: driver.find_element(By.XPATH, pathFacturas.replace('{contador2}', str(contador)).replace('{contador}', str(posFecha))).click()
                except: driver.find_element(By.XPATH, pathFacturas.replace('{contador2}', str(contador)).replace('{contador}', str(posFecha))).click()
                sleep(4)
                fechaFactura = driver.find_element(By.XPATH, pathFacturas.replace('{contador2}', str(contador)).replace('{contador}', str(posFecha)))
                fechaFactura = fechaFactura.get_attribute('title')
                print(f'Factura: {fechaFactura} Numero: {str(contador)}')

                fechaFactura = datetime.strptime(fechaFactura, '%d/%m/%Y')
                
                # Validacion rango 3 meses fecha obtenida
                if rango3Meses <= fechaFactura <= fechaHoy or fechaFactura >= fechaHoy:
                    print('Validando Cargo')
                    resultado = deteccion_cargo(driver)
                    if resultado != False: 
                        driver.find_element(By.XPATH, f'//a[contains(text(), "Pantalla Única de Consulta:")]').click()
                        return True, obtener_nombre_mes(fechaFactura.month), '-', resultado
                    else: contador += 1


                else: return False, '', 'No aplica: item facturacion', ''


            except: 
                try: 
                    driver.find_element(By.XPATH, pathSigFactura).click()
                    contador = 2
                except: return False, '', 'No aplica: item facturacion', ''


        # Obtencion posicion campo Fecha Limite de la Factura

        # posFechaLimite = obtencionColumna(driver, 'Fecha límite', pathEncabezadoFacturas)
        
        # # Busqueda y descarga de Facturas

        
        # buscandoFacturas = True
        # facturasDescargadas = []
        # fechasValidadas = []
        # contador = 2
        # while buscandoFacturas:
        #     try:

        #         # Obtencion Fecha de Factura
        #         fechaFactura = driver.find_element(By.XPATH, pathFacturas.replace('{contador2}', str(contador)).replace('{contador}', str(posFechaLimite)))
        #         fechaFactura = fechaFactura.get_attribute('title')
        #         print(fechaFactura)

        #         # Validacion Fecha Unica
        #         if fechaFactura not in fechasValidadas: 
        #             fechasValidadas.append(fechaFactura)
        #             fechaFactura = datetime.strptime(fechaFactura, '%d/%m/%Y')

        #             # Validacion rango 3 meses fecha obtenida
        #             if rango3Meses <= fechaFactura <= fechaHoy or fechaFactura >= fechaHoy:
        #                 driver.find_element(By.XPATH, pathFacturas.replace('{contador2}', str(contador)).replace('{contador}', str(posFechaLimite))).click()
        #                 sleep(1)
        #                 print('Esperando Cuadro Factura')

        #                 driver.find_element(By.XPATH, "//button[@title='Ver Factura']").click()

        #                 # Inicio de proceso de descarga de factura
        #                 cargandoFactura = True
        #                 while cargandoFactura:
        #                     try:
        #                         print('Descargando Factura')
        #                         # Cuadro Emergente con la factura
        #                         # driver.find_element(By.XPATH, pathCuadroPDF)
        #                         driver.find_element(By.XPATH, f'//span[contains(text(), "Visor de Factura")]')
        #                         sleep(7)

        #                         # Obtencion del url para descargar factura
        #                         # url = driver.find_element(By.XPATH, pathEmbedPDF)
        #                         url = driver.find_element(By.XPATH, "//embed[@id='archivoPDF']")
        #                         url = url.get_attribute('src')
        #                         print(url)
                                
        #                         if 'blob' in url:

        #                             # Inyeccion de codigo para descarga de pdf
        #                             driver.execute_script("""
        #                                 // arguments[0] es el blob que se pasa
        #                                 const blobUrl = arguments[0];
        #                                 const a = document.createElement('a');
        #                                 a.style.display = 'none';
        #                                 a.href = blobUrl;
        #                                 a.download = 'documento.pdf';
        #                                 document.body.appendChild(a);
        #                                 a.click()
        #                                 document.body.removeChild(a);
        #                             """, url)

        #                             # Inicio de validacion de pdf unico descargado
        #                             sleep(5)
        #                             descargandoPDF = True
        #                             contadorDescarga = 0
        #                             while descargandoPDF:

        #                                 ruta = Path(r'C:\Users\sistemas\Downloads')
        #                                 # ruta = Path(r'C:\Users\AjusteSucursales\Downloads')
        #                                 pdfs = list(ruta.glob('*.pdf'))
        #                                 if not pdfs:
        #                                     contadorDescarga += 1
        #                                     sleep(1)
        #                                     if contadorDescarga == 20: return False, None, 'Inconsistencia Siebel: PDF No descarga', '-'
        #                                 else:
        #                                     reciente = max(pdfs, key=lambda f: f.stat().st_mtime)
        #                                     if reciente not in facturasDescargadas:
        #                                         print('PDF MAS RECIENTE')
        #                                         facturasDescargadas.append(reciente)
        #                                         resultado, monto = extraer_texto_de_pdf_con_imagenes(reciente)
        #                                         os.remove(reciente)
        #                                         if resultado == True:
        #                                             descargandoPDF = False
        #                                             cargandoFactura = False
        #                                             buscandoFacturas = False
        #                                             fechaFactura = fechaFactura.strftime('%d/%m/%Y')
        #                                             print('→ Fin de extraccion de facturas')
        #                                             try: driver.find_element(By.XPATH, pathCierrePDF).click()
        #                                             except: pg.press('ESC')
        #                                             return True, fechaFactura,'', monto
        #                                         # elif len(facturasDescargadas) == 3 and resultado == False: return False, None, 'No Aplica: Item Facturacion', '-'
        #                                         else:
        #                                             try: driver.find_element(By.XPATH, pathCierrePDF).click()
        #                                             except: pg.press('ESC')
        #                                             sleep(5)
        #                                             descargandoPDF = False
        #                                             cargandoFactura = False
        #                                             contador += 1
        #                                     else:
        #                                         sleep(1)
        #                                         contadorDescarga += 1
        #                                         if contadorDescarga == 20: 
        #                                             descargandoPDF = False
        #                                             cargandoFactura = False
        #                                             contador += 1
                                
        #                         try: driver.find_element(By.XPATH, pathCierrePDF).click()
        #                         except: pg.press('ESC')

        #                         sleep(7)
        #                     except: 
        #                         sleep(1)
        #                         print('Esperando ventana PDF')
        #             # elif fechaFactura > fechaHoy: contador += 1
        #             else:
        #                 print('Fin de busqueda de Facturas')
        #                 buscandoFacturas = False
        #                 return False, None, 'No Aplica: Item Facturacion', '-'

                
        #         else: contador += 1

        #     except Exception as e:
        #         try: 
        #             driver.find_element(By.XPATH, "//td[@id='next_pager_s_3_l']").click()
        #             contador = 2
        #             sleep(7)
        #         except: 
        #             buscandoFacturas = False
        #             return False, None, 'No Aplica: Item Facturacion', '-'
        
 

    except Exception as e:
        print(f'ERROR ajustando el cargo extemporaneo. Caso NO. {no_caso}. CUENTA: {no_cuenta}')
        error = e
        print('Error ',error)
        return False, '',error

def aplicacionAjuste(driver, montoAjuste, mesFactura, noCN):

    try: 

        pathColumnasBAjuste2 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
        pathColumnasBAjuste = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
        pathColumnasBAInput = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input[2]'
        pathColumnasBAInput2 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input'
        pathColumnasBAInput3 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input'

        print('Generando Ajuste')
        ajusteCarga, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Solicitudes de Ajuste Applet de lista:Nuevo')
        if ajusteCarga == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga', '-'
        
        ajusteCarga, resultado = cargandoElemento(driver, 'input', 'aria-label', 'Importe del ajuste')
        if ajusteCarga == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga', '-'
        print('♥ Generando Ajuste ♥')
        
        # Importe Ajuste
        input_importe = driver.find_element(By.XPATH, "//input[@aria-label='Importe del ajuste']")
        input_importe.click()
        sleep(1)
        input_importe.clear()
        input_importe.send_keys(str(montoAjuste))
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
        input_motivo_ajuste.send_keys('CARGO POR PAGO EXTEMPORANEO')
        input_motivo_ajuste.send_keys(Keys.RETURN)
        print('♦ Motivo Ajuste Ingresado ♦')

        input_comentario = driver.find_element(By.XPATH, "//textarea[@aria-label='Comentarios']")
        input_comentario.click()
        sleep(1)
        input_comentario.send_keys(f'''SE APLICA AJUSTE POR CARGO EXTEMPORANEO FACTURA DEL MES DE {mesFactura} CN {noCN} ROBOT''')
        print('♦ Comentario Ingresado ♦')

        numeroAjuste = driver.find_element(By.XPATH, "//input[@aria-label='Número de Ajuste']")
        numeroAjuste = numeroAjuste.get_attribute("value")
        print(f'Ajuste Generado: {numeroAjuste}')

        driver.find_element(By.XPATH, "//button[@aria-label='Solicitud de Ajuste Applet de formulario:Guardar']").click()
        print('♦ Ajuste Guardado ♦')
        sleep(5)
        
        try:
            alert = Alert(driver)
            alert_txt = alert.text
            print(f'♦ {alert_txt} ♦')
            alert.accept()
            return False, f'Inconsistencia Siebel: {alert_txt}', numeroAjuste
        except: pass
        
        # Envio Ajuste
        print('♥ Enviando Ajuste ♥')
        ajusteCarga, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Solicitudes de Ajuste Applet de lista:Consulta')
        if ajusteCarga == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga', numeroAjuste

        # sleep(10000)
        posicion = obtencionColumna(driver, 'Número de Ajuste', pathColumnasBAjuste, pathColumnasBAjuste2)
        if posicion == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga', numeroAjuste
        
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
        if ajusteCarga == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga', numeroAjuste
        # sleep(10000)

        ajusteCarga, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Enviar Ajuste Applet de formulario:Aceptar')
        if ajusteCarga == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga', numeroAjuste
        print('♦ Ajuste Enviado ♦')
        sleep(5)
        return True, 'Cerrado', numeroAjuste

    except: pass

def cierreCancelacionCasoActividad(driver, no_cuenta, mesFactura, noCN, cancelacion = False):
    try: 

        try: driver.find_element(By.XPATH, "//a[@title='Pantalla Única de Consulta']").click()
        except: return False, 'Inconsistencia Siebel: Pantalla NO Carga'
        
        # Buscando Elemento
        lupa_busqueda_cuenta, resultado = cargandoElemento(driver, 'button', 'title', 'Pantalla Única de Consulta Applet de formulario:Consulta')
        if lupa_busqueda_cuenta == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(5)

        # Ingreso Cuenta
        input_busqueda_cuenta = driver.find_element(By.XPATH, "//input[@aria-label='Numero Cuenta']")
        input_busqueda_cuenta.click()
        sleep(1)
        input_busqueda_cuenta.send_keys(no_cuenta)
        input_busqueda_cuenta.send_keys(Keys.RETURN)
        print('♦ Cuenta Ingresada ♦')

        lupa_busqueda_cn, resultado = cargandoElemento(driver, 'input', 'aria-label', 'Perfil de Pago')
        if lupa_busqueda_cn == False: return False, 'Pendiente'
        print('♥ Cuenta OK! ♥')
        sleep(5)
        
        pathColumnasBActividad = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
        pathColumnasBAInput = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input[2]'

        pathOpcEstadosAc = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[5]/span'
        pathOpcEstadosAc2 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[4]/span'


        pathColumnasBCN = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
        pathColumnasBCNInput = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input[2]'
        pathColumnasBCNInput2 = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input'

        ajusteCarga, resultado = cargandoElemento(driver, 'button', 'aria-label', 'Casos de Negocio Applet de lista:Consulta')
        if ajusteCarga == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(5)

        # Ingreso CN
        print('→ Cierre Actividad')
        posicion = obtencionColumna(driver, 'Caso de Negocio', pathColumnasBCN)
        if posicion == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(3)


        try: input_busqueda_cn = driver.find_element(By.XPATH, pathColumnasBCNInput.replace('{contador}', posicion))
        except: input_busqueda_cn = driver.find_element(By.XPATH, pathColumnasBCNInput2.replace('{contador}', posicion))
        input_busqueda_cn.click()
        input_busqueda_cn.send_keys(noCN)
        input_busqueda_cn.send_keys(Keys.RETURN)
        print('→ CN Ingresado OK!')
        sleep(15)

        posicionEstado = obtencionColumna(driver, 'Estado', pathColumnasBCN)
        if posicionEstado == False: return False, 'Inconsistencia Siebel: Pantalla NO Carga'
        sleep(3)


        try: driver.find_element(By.XPATH, pathColumnasBCNInput.replace('{contador}', posicionEstado).replace('/input[2]', '')).click()
        except: driver.find_element(By.XPATH, pathColumnasBCNInput2.replace('{contador}', posicionEstado).replace('/input', '')).click()
        sleep(1)

        resultado, resultado2 = cargandoElemento(driver, '', '', '', f'//a[contains(text(), "{noCN}")]')
        if resultado != True: return False, 'Inconsistencia Siebel: Pantalla CN NO Carga'
        print('♦ CN Abierto ♦')
        # Actividades 1216628040569 1216701545980
        resultado, resultado2 = cargandoElemento(driver, '', '', '', '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/ul/li[2]/a')
        if resultado != True: return False, 'Inconsistencia Siebel: Pantalla CN NO Carga'
        print('♦ Actividades Abiertas ♦')
        sleep(10)

        # Comentarios
        posicionCA = obtencionColumna(driver, 'Comentarios', pathColumnasBActividad)
        if posicionCA == False: return False, 'Inconsistencia Siebel: Pantalla CN NO Carga'

        # Motivo Cierre

        if cancelacion == True:
            posicionMCA = obtencionColumna(driver, 'Motivo de la cancelación', pathColumnasBActividad)
            if posicionMCA == False: return False, 'Inconsistencia Siebel: Pantalla CN NO Carga'
        else:
            posicionMCA = obtencionColumna(driver, 'Motivo del cierre', pathColumnasBActividad)
            if posicionMCA == False: return False, 'Inconsistencia Siebel: Pantalla CN NO Carga'

        # Estado Actividad
        posicionEA = obtencionColumna(driver, 'Estado', pathColumnasBActividad)
        if posicionEA == False: return False, 'Inconsistencia Siebel: Pantalla CN NO Carga'

        
        # Ingres Comentarios

        driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicionCA).replace('/input[2]', '')).click()
        sleep(1)
        input_busqueda_comentarios = driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicionCA).replace('/input[2]', '/textarea'))
        try: input_busqueda_comentarios.click()
        except: return False, 'Inconsistencia Siebel: Asignacion'

        if cancelacion == True: input_busqueda_comentarios.send_keys(f'''SE CANCELA CN {noCN} POR FALTA DE INFORMACION FALTA DE SOPORTE. ROBOT''')
        else: input_busqueda_comentarios.send_keys(f'''SE APLICA AJUSTE POR CARGO EXTEMPORANEO FACTURA DEL MES DE {mesFactura} CN {noCN} ROBOT''')
        print('♥ Comentario Actividad Ingresado ♥')

        
        sleep(2)

        # Ingreso Motivo Cierre

        

        try:
            driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicionMCA).replace('/input[2]', '')).click()
            input_busqueda_mcierre = driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicionMCA))
            if cancelacion == True: input_busqueda_mcierre.send_keys('FALTA SOPORTE')
            else:  input_busqueda_mcierre.send_keys('SE APLICA AJUSTE')
            input_busqueda_mcierre.send_keys(Keys.RETURN)
            print('♥ Motivo Cierre Actividad Ingresado ♥')
        except Exception:
            try:
                # driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicion).replace('/input[2]', '')).click()
                input_busqueda_mcierre = driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicionMCA).replace('/input[2]', '/input'))
                if cancelacion == True: input_busqueda_mcierre.send_keys('FALTA SOPORTE')
                else:  input_busqueda_mcierre.send_keys('SE APLICA AJUSTE')
                input_busqueda_mcierre.send_keys(Keys.RETURN)
                print('♥ Motivo Cierre Actividad Ingresado ♥')
            except Exception as e: print(f'{str(e)}');  return False, 'Inconsistencia Siebel: Cambio Elementos'

        # Ingreso Estado

        driver.find_element(By.XPATH, pathColumnasBAInput.replace('{contador}', posicionEA).replace('/input[2]', '')).click()
        sleep(2)

        try: driver.find_element(By.XPATH, pathOpcEstadosAc).click()
        except:
            try: driver.find_element(By.XPATH, pathOpcEstadosAc2).click()
            except Exception as e: print(str(e)); return False, 'Inconsistencia Siebel: Cambio Elementos'
        
        if cancelacion == True: resultado, resultado2 = cargandoElemento(driver, '', '', '', '//div[contains(text(), "CANCELADA")]')
        else:  resultado, resultado2 = cargandoElemento(driver, '', '', '', '//div[contains(text(), "CERRADA")]')
        if resultado != True: return False, 'Inconsistencia Siebel: Pantalla NO Carga'

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
            else: return False, f'Inconsistencia Siebel: {alert_txt}'

        except: pass
        sleep(5)

        print('♦ Cuenta terminada ♦')
        return True, 'Cerrado'


    except Exception as e: print(e); sleep(10000)