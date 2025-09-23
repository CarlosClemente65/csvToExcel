# DsecsvToExcel v2.0
## Programa para generar un fichero Excel a partir de un CSV

### Desarrollado por Carlos Clemente (02-2025)

### Control de versiones
- Version 1.0.0 - Primera version funcional
- Version 1.1.0 - Añadida posibilidad de agrupar en un solo fichero excel varias hojas
- Version 1.2.0 - Corregido problema de ejecucion con muchas lineas en el origen
- Version 1.2.1 - Modificada carpeta de salida de librerias
- Version 1.2.2 - Añadido metodo para convertir la plantilla si viene como Excel 97-2003 (xsl) a Excel 2007 (xlsx)
- Version 2.0.0 - Modificada la ejecucion para pasar una clave y un guion con los parametros necesarios
- Version 2.1.0 - Modificado la lectura del .csv para controlar la codificacion, ajustes en varios procesos y que la salida sea siempre en .xlsx

Instrucciones:
- Se debe pasar como parametro la clave de ejecucion seguido con el nombre de un fichero que se usara como 'guion' con los parametros de ejecucion
- En el fichero 'guion.txt' se incluiran los parametros con el formato 'clave=valor'
- Se pueden establecer la hoja y la celda en la que se insertarán los datos (por defecto hoja 1 y celda A1)
- Se puede utilizar un libro personalizado como plantilla
- Si se procesan varios ficheros con el mismo nombre de 'salida' se añaden hojas al mismo libro (parametro agrupar = SI)
- El parametro 'AGRUPAR='SI (defecto NO). Añade hojas al final del fichero de salida (nombre de la primera mas un contador) o borra (valor 'NO') el fichero de salida previamente
- El parametro 'INSERTAR=NO' (defecto SI). Copia los datos en el fichero segun la hoja pasada como parametro o añade hojas nuevas al final del fichero (valor 'SI')
- Los campos del CSV deben separarse con punto y coma (campo1;campo2;campoN)
- Permite añadir formulas al CSV teniendo en cuenta lo siguiente:
	* La formula debe comenzar por '#F#' (sin las comillas) para evitar errores en la transformacion; luego se transforma en formula
	* La separacion de parametros deben hacerse con una coma en vez del punto y coma
	* El nombre de las funciones debe hacerse en ingles
	* Ejemplo de formula (generar un hipervinculo a un fichero): 
		#F#HYPERLINK("C:\DOCUMENTOS\000003480.PDF","000003480.PDF")
- Se genera el fichero "respuesta.txt" con el resultado de la operacion
	
- Uso de la aplicacion:
	* cstToExcel.exe clave guion.txt
- Opciones del guion.txt:
	* ENTRADA=archivo.csv (obligatorio)
	* SALIDA=archivo.xlsx (obligatorio)
	* PLANTILLA=plantilla.xlsx (opcional)
	* CELDA=A1 (defecto)
	* HOJA=1 (defecto)
	* AGRUPAR=SI (defecto NO)
	* INSERTAR=NO (defecto SI)
	
