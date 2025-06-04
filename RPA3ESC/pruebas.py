from funcionalidad import *
from login import *
from datetime import datetime, timedelta, date
import autoit as it
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import apiCyberHubOrdenes as api



  
def pruebas(driver, cuenta, cn, propietario):

    nameEncabezadosActividades = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
    inputBusquedaActividades = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/input'


    try: driver.find_element(By.XPATH, "//a[@title='Actividades']").click()
    except: return 'Inconsistencia Siebel: Elemento Actividades', '', '-'

    # Buscando Elemento
    lupa_busqueda_cuenta = cargandoElemento(driver, 'select', 'title', 'Visibilidad')
    if lupa_busqueda_cuenta == False: return 'Inconsistencia Siebel: Pantalla NO Carga', '-', '-'
    sleep(5)

    # Selecciona Reasignacion Manual
    resultado = cargandoElemento(driver, '', '', '', f'//option[contains(text(), "Reasignación manual de actividades")]')
    if resultado != True: return 'Inconsistencia Siebel: Pantalla NO Carga'
    
    # Buscando Elemento
    lupa_busqueda_cuenta = cargandoElemento(driver, 'button', 'title', 'Reasignar propietario Applet de lista:Consulta')
    if lupa_busqueda_cuenta == False: return 'Inconsistencia Siebel: Pantalla NO Carga', '-', '-'
    sleep(5)


    posicionCN = obtencionColumna(driver, 'Nº de caso de negocio', nameEncabezadosActividades)
    if posicionCN == False: return 'Inconsistencia Siebel: Pantalla NO Carga'

    driver.find_element(By.XPATH, inputBusquedaActividades.replace('{contador}', posicionCN).replace('/input', '')).click()
    input_busqueda_cn_act = driver.find_element(By.XPATH, inputBusquedaActividades.replace('{contador}', posicionCN))
    input_busqueda_cn_act.send_keys(cn)
    input_busqueda_cn_act.send_keys(Keys.RETURN)
    sleep(15)

    posicion = obtencionColumna(driver, 'Propietario', nameEncabezadosActividades)
    if posicion == False: return 'Inconsistencia Siebel: Pantalla NO Carga'

    driver.find_element(By.XPATH, inputBusquedaActividades.replace('{contador}', posicion).replace('/input', '')).click()
    input_busqueda_prop_act = driver.find_element(By.XPATH, inputBusquedaActividades.replace('{contador}', posicion))
    # input_busqueda_prop_act.clear()
    sleep(5)
    input_busqueda_prop_act.send_keys(propietario)
    
    sleep(10000)
    sleep(15)
    driver.find_element(By.XPATH, inputBusquedaActividades.replace('{contador}', posicionCN).replace('/input', '')).click()
    print('♦ CN Abierto ♦')

# fechaHoy = '12/12/2018'
# fechaHoy = datetime.strptime(fechaHoy, '%d/%m/%Y').date()
# print(fechaHoy)

# p = '26/03/2025 14:09:00'
# p2 = '26/03/2025 15:09:00'
# p = datetime.strptime(p, '%d/%m/%Y %H:%M:%S')
# p2 = datetime.strptime(p2, '%d/%m/%Y %H:%M:%S')
# if p2 > p: print('aqui')

driver, _ = loginSiebel('ftorresma', 'Tobias1949++')
sleep(15)
# sleep(10000)
inicio(driver, '55608277', '1214125618042', 'retencion 0', sinEquipo = False)
# p = inicio(driver, '35848611', '1210667240142', 'Retencion', True)
# print(f'→ {p} ←')
# p = cierreActividad(driver, '35848611', '1210667240142', 'Retencion')
# print(f'→ {p} ←')
# sleep(10000)
# p = pruebas(driver, '70503926', '1208869478715', 'CVSUC9500410')
# print(p)




