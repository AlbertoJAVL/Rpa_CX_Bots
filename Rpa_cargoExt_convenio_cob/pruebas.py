from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from funcionalidad import validacionTipoCN, validacionCargoExt, busqueda_factura, aplicacionAjuste, cierreCancelacionCasoActividad
from logueo import login_siebel
from time import sleep

import base64
import time
from datetime import datetime

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

def prueba(driver):

    while True:
        try:
            sleep(10)
            driver.find_element(By.XPATH, "/html/body/div[21]/div[2]")
            url = driver.find_element(By.XPATH, "/html/body/div[21]/div[2]/embed")
            url = url.get_attribute('src')
            print(url)
            if 'blob' in url:
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
        except: print('intentando')


    embed = WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'embed[type="application/x-google-chrome-pdf"]')))
    pdf_base64 = driver.execute_async_script("""
        const callback = arguments[arguments.lenght -1];
        // Selecciona el embed
        const embed = document.querySelector('embed[type="application/x-google-chrome-pdf"]');
        // Toma la URL sin fragmento (#toolbar)
        const blobUrl = (embed.getAttribute('original-url') || embed.src).split('#')[0];
        fetch(blobUrl)
            .then(resp => resp.arrayBuffer())
            .then(buffer => {
                let binary = '';
                const bytes = new Uint8Array(buffer);
                for (let = i; i < bytes.byteLength; i++) {
                    binary += String.fromCharCode(bytes[i]);
                }
                // Devuelve el PDF en base64
                callback(btoa(binary));
            })
            .catch(err => {
                consele.error(err);
                callback(null);
            })
    """)

    if not pdf_base64:
        print('No se puede descar el pdf')

    pdf_bytes = base64.b64decode(pdf_base64)
    with open("documento_descargado.pdf", "wb") as f:
        f.write(pdf_bytes)

    print('Archivo descargado OK!')
    sleep(10000)

# p = '10/10/2025'
# p = datetime.strptime(p, '%d/%m/%Y')
# mes = p.month
# print(obtener_nombre_mes(mes))
driver, status_logueo = login_siebel('apiliado', 'Fabuloso#2025')
cierreCancelacionCasoActividad(driver, '34749543', 'abril', '1216701794419', False)
sleep(100000)
busqueda_factura(driver, '38104716', '1214746232721')
# # prueba(driver)
# # validacionTipoCN(driver, '1214746232721')
# # validacionCargoExt(driver, '38104716', '1214746232721')
# # resultado, fechaFactura, error, montoAjuste = busqueda_factura(driver, '38104716', '')
# # aplicacionAjuste(driver, montoAjuste, fechaFactura[3:5], '1214746232721')

# cierreCancelacionCasoActividad(driver, '61541911', '04','1214738158422')