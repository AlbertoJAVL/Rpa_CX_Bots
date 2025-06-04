import json

fechasInput = [
        '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[6]'
]

home = {
        'actividades': {
                'id': '',
                'name': '',
                'xpath': '/html/body/div[1]/div/div[3]/div/div/div[1]/div[3]/ul/li[4]',
        },
        'pantalla_unica':{
                                'id': '',
                                'name': '',
                                'xpath': '//*[@id="s_sctrl_tabScreen"]/ul/li[2]',
        },
        'home_from_sidebar': {
                'id': 'ui-id-159',
                'name': '',
                'xpath': '/html/body/div[1]/div/div[3]/div/div/div[1]/div[3]/ul/li[1]/a',
                
        },
        'pantalla_casos':{
                'id': '',
                'name': '',
                'xpath': '/html/body/div[1]/div/div[3]/div/div/div[1]/div[3]/ul/li[8]/a',
                },
        }
pantalla_casos_negocio = {
                        'lupa':{
                                'id':'',
                                'name':'s_2_1_13_0',
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[1]/div[2]/button[2]'
                        },
                        'ingresar_caso':{
                                'id':'',
                                'name':'SR_Number',
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/input'
                        },
                        'motivo_cliente':{
                                'id':'',
                                'name':'TT_Motivo_Cliente',
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[26]/input'
                        },
                        
                        'barra_caso_negocio':{
                                'id':'',
                                'name':'SR_Number',
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/input'
                        },
                        }

pantalla_unica = {
                'lupa': {
                        'id': 's_12_1_148_0_Ctrl',
                        'name': 's_12_1_148_0',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/form/div/span/div[1]/div[3]/button[6]',
                        },      
                'lupa_caso_negocio': {
                                'id': '',
                                'name': 's_1_1_18_0',
                                'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/div/div/form/span/div/div[1]/div[2]/button[2]'
                                         
                        },
                'ingresar_cuenta':{
                        'id': '',
                        'name': 's_12_1_154_0',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/form/div/span/div[3]/div/div/table/tbody/tr[3]/td[4]/div/input',
                        },
                'saldo_pendiente': {
                        'id': '',
                        'name': 's_12_1_32_0',
                        'xpath' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/div/div/form/div/span/div[3]/div/div/table/tbody/tr[5]/td[7]/div/input',
                        'xpath2': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/form/div/span/div[3]/div/div/table/tbody/tr[7]/td[7]/div/input',
                        },         
                'inputCuenta' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/div/div/form/div/span/div[3]/div/div/table/tbody/tr[3]/td[4]',
                'subtipo_cuenta':{
                                'id': '',
                                'name':'s_12_1_65_0',
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/div/div/form/div/span/div[3]/div/div/table/tbody/tr[7]/td[9]/div/input',
                                'xpath2' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/form/div/span/div[3]/div/div/table/tbody/tr[7]/td[9]/div/input'
                },
                'tipo_cuenta':{
                                'id': '',
                                'name':'s_12_1_193_0',
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/div/div/form/div/span/div[3]/div/div/table/tbody/tr[5]/td[9]/div/input',
                                'xpath2':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/form/div/span/div[3]/div/div/table/tbody/tr[5]/td[9]/div/input'
                },
                 'tipo_cuenta2':{
                                'id': '',
                                'name':'s_12_1_193_0',
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/div/div/form/div/span/div[3]/div/div/table/tbody/tr[5]/td[9]/div/input'
                },
                'antiguedad':{
                                'id': '',
                                'name':'s_12_1_36_0', 
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/div/div/form/div/span/div[3]/div/div/table/tbody/tr[13]/td[3]/div/input'
                },
                'estado_cuenta':{
                                'id': '',
                                'name':'s_12_1_195_0',
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/div/div/form/div/span/div[3]/div/div/table/tbody/tr[9]/td[9]/div/input',
                                'xpath2' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/form/div/span/div[3]/div/div/table/tbody/tr[9]/td[9]/div/input'
                },                
                'sub_estado_cuenta':{
                                'id': '',
                                'name':'s_12_1_111_0',
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/div/div/form/div/span/div[3]/div/div/table/tbody/tr[11]/td[9]/div/input'
                },

                'ingresar_caso':{
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/input'
                },
                'historial_facturas':{
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[1]'
                },
                'factura_actual':{
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/a'
                },
                'detalles': {
                        'xpath':'//*[@id="a_2"]/div[1]'
                },                
                'profundiza_caso': {
                        'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/a]'
                },
                'profundiza_actividades':{
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/ul/li[2]'
                },
                'profundiza_actividad_id': {
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[1]'
                },
                'estado_caso':{
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/div/span/div[3]/div/div/table/tbody/tr[3]/td[8]/div/input'
                },
                'saldos' : {
                        'xpath1':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/div/div/form/div/span/div[3]/div/div/table/tbody/tr[7]/td[7]/div/span',
                        'xpath2':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/form/div/span/div[3]/div/div/table/tbody/tr[7]/td[7]/div/span',
                        'saldoMes1' : '/html/body/div[21]/div[2]/div/div/div/table[2]/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[2]/span/input',
                        'saldoMes2' : '/html/body/div[11]/div[2]/div/div/div/table[2]/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[2]/span/input',
                        'saldo301' : '/html/body/div[21]/div[2]/div/div/div/table[2]/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td[2]/span/input',
                        'saldo302' : '/html/body/div[11]/div[2]/div/div/div/table[2]/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td[2]/span/input',
                        'saldo601' : '/html/body/div[21]/div[2]/div/div/div/table[2]/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/span/input',
                        'saldo602' : '/html/body/div[11]/div[2]/div/div/div/table[2]/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/span/input',

                },
                'lineaNegocio' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/div/div/form/div/span/div[3]/div/div/table/tbody/tr[7]/td[3]/div/input'
        }

actividades = {
                'menu_actividades': {
                                'id': '',
                                'name': '',
                                'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[1]/div[1]/select',
                },
                'mis_actividades': {
                                'id': '',
                                'name': '',
                                'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[1]/div[1]/select/option[1]',
                },
                'busqueda_actividades':{
                                'id': '',
                                'name': 's_2_1_0_0',
                                'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[2]/div[2]/button',
                },
                'no_cuenta':{
                                'id': '1_Account_Name',
                                'name': 'Account_Name',
                                'xpath': '//*[@id="1_s_2_l_Account_Name"]',
                },                
                'no_caso':{
                                'id': '1_s_2_l_SR_Number',
                                'name': 'SR_Number',
                                'xpath': '//*[@id="1_s_2_l_SR_Number"]',
                },
                'inicio':{
                        'xpath': ''
                }
        



}

facturas = {
        'fecha_inicial' : {
                        'id': '1_PeriodStartDate',
                        'name': 'PeriodStartDate',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[9]/input'
        }    
}

solicitud_ajuste =  {

                'lupa' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[1]/div[2]/button[2]',

                'engrane' : {
                                'id':'',
                                'name':'',
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[1]/div[2]/span[1]/span/button'
                        },
                
                'opcionOrdenar' : {
                                'id':'',
                                'name':'',
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[1]/div[2]/span[1]/span/ul/li[34]'
                        },

                'ordenarPor' : {
                                'id':'',
                                'name':'',
                                'xpath':'/html/body/div[21]/div[2]/div/div/div/form/table/tbody/tr[1]/td/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input'
                        }, 

                'descenciente' : {
                                'id':'',
                                'name':'',
                                'xpath':'/html/body/div[21]/div[2]/div/div/div/form/table/tbody/tr[1]/td/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/table/tbody/tr/td[3]/input[2]'
                        }, 
                
                'aceptarBTN' : {
                                'id':'',
                                'name':'',
                                'xpath':'/html/body/div[21]/div[2]/div/div/div/form/table/tbody/tr[2]/td/table/tbody/tr/td/span[1]/button'
                        },
                
                'elemMotivo' : {
                                'id':'',
                                'name':'',
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[7]/input'
                        },

                'motivoInput' : {
                                'id':'',
                                'name':'',
                                'xpath':'/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[7]/input'
                        },

                'nueva_solicitud' : {
                                'id': '1_PeriodStartDate',
                                'name': 's_5_1_7_0',
                                'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[1]/div[2]/button[1]'
                        },
                'importe':  {
                        'name':'',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/div/table[2]/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr[2]/td[2]/span/input'
                        },
                'guardar':{
                        'name':'s_30_1_17_0',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/div/table[1]/tbody/tr/td[1]/span/div/div[3]/button[1]'
                },
                'aFavor' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/div/table[2]/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/span/input',
                'mAjuste' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/div/table[2]/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr[4]/td[2]/span/input',
                'comentario' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/div/table[2]/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]',
                'cancelar':{
                        'name':'s_30_1_16_0',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/div/table[1]/tbody/tr/td[1]/span/div/div[3]/button[2]'
                },
                'consulta_saldos':{
                        'name':'s_12_1_24_0',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/div/div/div/form/div/span/div[1]/div[3]/button[8]'
                }         
        }

historial_ajustes = {

        'engrane' : {
                        'id': '',
                        'name': '',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[4]/div/form/span/div/div[1]/div[2]/span[1]/span/button',
                },

        'opcionOrdenar' : {
                        'id': '',
                        'name': '',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[4]/div/form/span/div/div[1]/div[2]/span[1]/span/ul/li[34]',
                },

        'ordenarPor' : {
                        'id': '',
                        'name': '',
                        'xpath': '/html/body/div[21]/div[2]/div/div/div/form/table/tbody/tr[1]/td/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]',
                },
        
        'descendiente' : {
                        'id': '',
                        'name': '',
                        'xpath': '/html/body/div[21]/div[2]/div/div/div/form/table/tbody/tr[1]/td/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/table/tbody/tr/td[3]/input[2]',
                },
        
        'aceptarBTN' : {
                        'id': '',
                        'name': '',
                        'xpath': '/html/body/div[21]/div[2]/div/div/div/form/table/tbody/tr[2]/td/table/tbody/tr/td/span[1]/button',
                },
        
        'inputFecha' : {
                        'id': '',
                        'name': '',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[4]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[2]',
                        
                }

}

caso_negocio = {

        'lupa' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[1]/div[2]/button[2]',
        'inputCN' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/input',
        'comentario' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[8]',
        'motivo' : {
                        'id': '',
                        'name': '',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[5]',
                },

        'subMotivo' : {
                        'id': '',
                        'name': '',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[6]',
                },

        'solucion' : {
                        'id': '',
                        'name': '',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[7]',
                },

        'motivoCliente' : {
                        'id': '',
                        'name': '',
                        'xpath': '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[9]',
                },


        'actividades' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/ul/li[2]/a',
        'columnasAct' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div',
        'filasAct' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]',
        'selecEstado' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]/span',
        'estado' : '/html/body/div[1]/div/div[5]/div/div[8]/ul[12]/li[{contador}]/div'
        
}

pantallaActividades = {
        'comentario' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[5]',

        'motivo' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[6]',

        'motivoCancelacion' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[18]',

        'estado' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[3]/div[1]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[7]',

        'estadoCN' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[1]/div/form/div/span/div[3]/div/div/table/tbody/tr[3]/td[8]/div/input'
}

historial_pago = {
        'engrane' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/div/div/form/span/div/div[1]/div[2]/span[1]/span/button',
        'opcionOrden' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/div/div/form/span/div/div[1]/div[2]/span[1]/span/ul/li[34]/a',
        'opcionColumnas' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/div/div/form/span/div/div[1]/div[2]/span[1]/span/ul/li[33]/a',
        'input' : '/html/body/div[21]/div[2]/div/div/div/form/table/tbody/tr[1]/td/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]',
        'orden' : '/html/body/div[21]/div[2]/div/div/div/form/table/tbody/tr[1]/td/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/table/tbody/tr/td[3]/input[2]',
        'btnAceptar' : '/html/body/div[21]/div[2]/div/div/div/form/table/tbody/tr[2]/td/table/tbody/tr/td/span[1]/button',
        'fechaPago' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[3]',
        'columnaAgregada' : '/html/body/div[10]/div[2]/div/div/div/form/table[1]/tbody/tr[2]/td[1]/select/option[2]',
        'agregarColumna' : '/html/body/div[10]/div[2]/div/div/div/form/table[1]/tbody/tr[2]/td[2]/span[1]/span/a',
        'guardarColumna' : '/html/body/div[21]/div[2]/div/div/div/form/table[2]/tbody/tr/td/span[2]/button',
        'monto' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[1]/div[3]/div/div/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[10]'
}

solicitudesAjustes = {
        'engrane' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[1]/div[2]/span[1]/span/button',
        'opcOrdenar' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[1]/div[2]/span[1]/span/ul/li[34]/a',
        'ordenarPor' : '/html/body/div[21]/div[2]/div/div/div/form/table/tbody/tr[1]/td/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/input',
        'descenciente' : '/html/body/div[21]/div[2]/div/div/div/form/table/tbody/tr[1]/td/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/table/tbody/tr/td[3]/input[2]',
        'btnAceptar' : '/html/body/div[21]/div[2]/div/div/div/form/table/tbody/tr[2]/td/table/tbody/tr/td/span[1]/button',
        'fechaAjuste' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]',
        'lupa' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[1]/div[2]/button[2]',
        'motivoAjuste' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]',
        'columnas' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
}

ajuste_CE = {
        'comentario' :  '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/div/table[2]/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/span/textarea',
        'motivo' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/div/table[2]/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr[4]/td[2]/span/input',
        'aplicar' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/div/table[2]/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/span/input',
        'monto' : '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/div/div/form/div/table[2]/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr[2]/td[2]/span/input'
}


profundizarCN = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[2]/a'
columnaCN = "/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div"
filasCN = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]'
columnaSA = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[2]/div/table/thead/tr/th[{contador}]/div'
filasSA = '/html/body/div[1]/div/div[5]/div/div[8]/div[2]/div[1]/div/div[2]/div[2]/div[3]/div/form/span/div/div[3]/div/div/div[3]/div[3]/div/div[2]/table/tbody/tr[2]/td[{contador}]'