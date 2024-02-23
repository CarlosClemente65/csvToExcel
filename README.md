# DsecsvToExcel v1.0
## Programa para generar un fichero Excel a partir de un CSV

### Desarrollado por Carlos Clemente (02-2024)

Instrucciones:
- Permite generar un libro de Excel a partir de un fichero CSV
- Se pueden establecer la hoja y la celda en la que se insertarán los datos; si no se indican se insertaran en la hoja 1 y celda A1
- Se puede utilizar un libro personalizado como plantilla
- Los campos del CSV deben separarse con punto y coma (campo1;campo2;campoN)
- Permite añadir formulas al CSV teniendo en cuenta lo siguiente:
	* El simbolo de igual se debe sustituir por '#F#' (sin las comillas) para evitar errores en la transformacion
	* La separacion de parametros de las formulas deben hacerse con una coma en vez del punto y coma
	* El nombre de las funciones debe hacerse en ingles
	* Ejemplo de formula (generar un hipervinculo a un fichero): #F#HYPERLINK("C:\DOCUMENTOS\000003480.PDF","000003480.PDF")
- Se genera el fichero "respuesta.txt" con el resultado de la operacion
	
- Parametros de ejecucion:
	* entrada = archivo.csv (obligatorio)
	* salida = archivo.xlsx (obligatorio)
	* plantilla = plantilla.xlsx (opcional)
	* celda = A1 (defecto)
	* hoja = 1 (defecto)
	
