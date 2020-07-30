'''
Creado : 2019-07-22
@autor: Pablo Akerman
Título: Generador de APs automático (AP = Ante Proyecto)
'''

from docxtpl import DocxTemplate
from datetime import date
import csv
import os

# Mesnsaje al usuario
print ('\nHola!!')
print ('Vamos a procesar todos los Proyectos que estén en el estado '
+ 'Elaborar AP')
print('La salida del CRM (exportar a CSV) tiene que estar en el archivo: '
+ 'reporte.csv')
print('Acordate de tener la PLANTILLA_AP.docx en el directorio \plantillas')
input('Estás preparado? (tocá cualquier tecla)')

dia_hoy = date.today()
dia_ar_dir = dia_hoy.strftime("%d_%m_%Y")	#formato nombre del directorio
dia_ar = dia_hoy.strftime("%d/%m/%Y")		#formato latino argentina (ar)

# TODO directorio = 'salida_' + dia_ar_dir +'_A'
# La A es para ir cambiando la letra, si es que ya existe el directorio.
# Mejor un contador. _01 _02 _03. Creo que confunde más el número.

# Creación del directorio de salida donde se guarda el word del AP consolidado
directorio = 'salida_' + dia_ar_dir
if not os.path.exists(directorio):
    os.mkdir(directorio)
    print(f'\nSe creó el directorio: {directorio}\n')
else:
    print(f'\nEl directorio {directorio} YA EXISTE. Te guardo los APs ahí.\n')

# generando el documento de salida
with open('reporte.csv', mode='r') as csv_file:
    reporte_csv = csv.DictReader(csv_file, delimiter=';')
    cantidad_lineas = 0
    for col in reporte_csv:
        if col["Actividad"] != 'Elaboracion de AP por Preventa':
            continue
        plantilla = DocxTemplate('plantillas/plantilla_ap.docx')
# TODO agregar un try para capturar los errores.
# Se le muestra al usuario los proyectos que se están procesando
        print(f'\t Nro CERES {col["Caso"]} de la empresa {col["Organizacion"]}'
        + f', título:{col["Titulo"]}.')
        print(f'\t \t contacto:{col["Contacto"]} Email: {col["Emailcontacto"]}')
        cantidad_lineas += 1
# Diccionario clave = etiqueta para plantilla : valor = columna reporte_csv
        contexto = {
            'nombre_cliente' : col["Organizacion"],
            'contacto_nombre' : col["Contacto"],
            'contacto_email' : col["Emailcontacto"],
            'contacto_tel' : "falta telefono",
            'nombre_proyecto' : col["Titulo"],
            'ejecutivo' : col["Ejecutivo cuentas"],
            'fecha' : dia_ar,
            'numero_CRM' : col["Caso"],
            'autor' : col["Preventa"],
            'central' : col["Localidad"],
            'central_provincia' : col["Provincia"],
            'remota' : col["Localidad remota"],
            'remota_provincia' : col["Provincia remota"],
            'descripcion'	:	col["Descripcion"],
			'servicio'	:	col["Servicio"],
        }

        plantilla.render(contexto)
        plantilla.save(directorio+'/AP_'+col["Caso"]+'.docx')
        contexto.clear()
print(f'\nSe procesaron {cantidad_lineas} de Proyectos.\n')
csv_file.close()
input('Press Enter to continue...')

"""
TODO
En el campo email eliminar < > por algún motivo rompe el word.
TODO
Generar el directorio \salida_dd_mm_aa
Verificar si existe -> \salida_dd_mm_aa_vN+1
TODO
try. ver cómo pongo que falta un campo!
Traceback (most recent call last):
  File "C:/Coding/Python/autoap/autoap.py", line 49, in <module>
    print(f'\t , contacto:{col["Contacto"]} Email: {col["Email Contacto"]}')
KeyError: 'Contacto'
"""
