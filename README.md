# AUTO AP
Pequeño script de automatización para la oficina.
El script toma la salida del CRM reporte.csv; que incluye todos los proyectos del usuario. Para los proyectos que están en estado "Elaborar AP", genera un documento AP Word a partir de una plantilla_ap.docx.
Se genera un directorio con la fecha del día de ejecución del script, donde se guardan los APs.

## AP
AP es la contracción de ante proyecto. Es un documento que contiene información del proyecto a ser implementado.

## Ejecución / Instalación
Se ejecuta como un script. Puede distribuirse a usuarios windows pasándolo como una aplicación. Utilicé pyinstaller:
>pyinstaller --onefile autoap.py

## TODO Pendientes
* En el campo email eliminar < > por algún motivo rompe el word.
* Generar el directorio \salida_dd_mm_aa (hecho)
* Verificar si existe directorio => \salida_dd_mm_aa_vN+1
