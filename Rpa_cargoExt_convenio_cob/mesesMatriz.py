import datetime

fecha = "05/06/2023 19:30:10"
fecha = datetime.datetime.strptime(fecha, '%d/%m/%Y %H:%M:%S')
fecha = fecha.strftime('%Y-%m-%d %H:%M:%S')
fecha = datetime.datetime.strptime(fecha, '%Y-%m-%d %H:%M:%S')
print(fecha)

fecha2 = datetime.datetime.now()
print(fecha2)


diferencia_fechas = fecha2 - fecha

tres_meses = datetime.timedelta(days=3*30)

if diferencia_fechas > tres_meses:
    print('El ultimo ajuste es mayor a 3 meses')
else:
    print('El ultimo ajuste es menor a 3 meses')