using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using Ude;
using Excel = Microsoft.Office.Interop.Excel;


namespace csvToExcel
{

    public static class Procesos
    {
        //Metodo para leer el CSV creando una lista de objetos que almacenara las lineas y dentro otra lista de objetos con cada uno de los campos
        public static List<List<object>> leerCSV(string archivoCSV)
        {
            string[] lineas; //Contiene las lineas del fichero

            try
            {
                //Deteccion de la codificacion del fichero csv
                var encodingDetector = new CharsetDetector();
                using(var filestream = new FileStream(archivoCSV, FileMode.Open))
                {
                    encodingDetector.Feed(filestream);
                    encodingDetector.DataEnd();
                }

                //Almacena la codificacion en la variable 'codificacion' para luego pasarla como parametro al leer el fichero y evitar que los caracteres como la 'Ñ' salgan con simbolos.
                var charset = encodingDetector.Charset;
                var encoding = charset != null ? Encoding.GetEncoding(charset) : Encoding.Default;

                //Lee todas las lineas del archivoCSV 
                using(StreamReader fichero = new StreamReader(archivoCSV, encoding))
                {
                    lineas = fichero.ReadToEnd().Split('\n');
                }

                //Almacena todos los campos de la linea en la lista de datos
                List<List<object>> datos = new List<List<object>>();

                for(int i = 0; i < lineas.Length; i++)
                {
                    if(!string.IsNullOrEmpty(lineas[i])) //Evita almacenar lineas vacias
                    {
                        //Divide cada linea en campos con el separador ';'. Ademas se establece cada campo con el tipo de valor que le corresponde segun su valor (int, float, DateTime, etc.
                        List<object> linea = lineas[i].Split(';').Select(x => tipoValor(x)).ToList<object>();
                        datos.Add(linea); //Se inserta en cada linea todos los campos en los que se haya dividido
                    }
                }

                return datos;
            }

            catch(Exception ex)
            {
                throw new Exception("Error al procesar el CSV" + ex);
            }
        }

        //Convierte cada objeto 'string' al tipo que le corresponde segun su valor.
        private static object tipoValor(string value)
        {
            //Intenta convertir a entero
            if(int.TryParse(value, out int intValue))
            {
                return intValue;
            }
            //Intenta convertir  decimal
            else if(decimal.TryParse(value, out decimal decimalValue))
            {
                return decimalValue;
            }
            //Intenta convertir a fecha
            else if(DateTime.TryParse(value, out DateTime dateTimeValue))
            {
                return dateTimeValue;
            }
            // Si no se puede convertir, se mantiene como string
            else
            {
                return value;
            }
        }


        public static void exportaXLSX(List<List<object>> datos)
        {
            //Asignacion de los valores de las variables
            string plantillaExcel = Program.plantillaExcel;
            int fila = Program.fila;
            int columna = Program.columna;
            int hoja = Program.hoja;
            string ficheroExcel = Program.ficheroExcel;

            string ficheroPlantilla = string.IsNullOrEmpty(plantillaExcel) ? "fichero.xlsx" : plantillaExcel;

            //Si la plantilla esta en formato Excel 97-2003 se convierte a Excel 2007
            MemoryStream ms;
            Stream fileStream;
            if(Path.GetExtension(ficheroPlantilla).Equals(".xls", StringComparison.OrdinalIgnoreCase))
            {
                ms = ConvertirAXlsx(ficheroPlantilla);
                fileStream = ms; //Carga la plantilla convertida en el fileStream
            }
            else
            {
                //Carga la plantilla original en el fileStream
                fileStream = new FileStream(ficheroPlantilla, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            }

            using(fileStream)
            {
                // Crea el libro con el contenido de la plantilla que corresponda
                XLWorkbook libroPlantilla = new XLWorkbook(fileStream); 

                //Asigna el tipo a la variable de la hoja 
                IXLWorksheet hojaPlantilla = null;
                try
                {
                    //Cargar la plantilla y si no existe crear un libro nuevo vacio
                    if(string.IsNullOrEmpty(plantillaExcel))
                    {
                        //libroPlantilla = new XLWorkbook(fileStream); //Si no se pasa la plantilla se crea un libro nuevo
                        hojaPlantilla = libroPlantilla.Worksheet(1); //Se crea la hoja 1 ya que no existe el libro
                    }
                    else
                    {
                        //Chequea que la hoja pasada existe en la plantilla para evitar una excepcion
                        if(hoja > libroPlantilla.Worksheets.Count)
                        {
                            Program.textoLog.AppendLine($"La hoja {hoja} no existe en la plantilla.");
                            return;
                        }
                        else
                        {
                            hojaPlantilla = libroPlantilla.Worksheet(hoja);
                        }
                    }

                }
                catch(Exception ex)
                {
                    Program.textoLog.AppendLine(ex.Message + $" Fichero = {ficheroExcel}. Hoja = {hoja}");
                    //throw new IndexOutOfRangeException(ex.Message + $" Fichero = {ficheroExcel}. Hoja = {hoja}");
                }

                try
                {
                    // Escribir los datos en la plantilla
                    for(int l = 0; l < datos.Count; l++) //Procesa las filas
                    {
                        for(int c = 0; c < datos[l].Count; c++) //Procesa las columnas
                        {
                            //Asigna el contenido de la celda segun la fila/columna procesada
                            object contenidoCelda = datos[l][c];

                            //Asigna la celda en la que grabar el valor
                            var cell = hojaPlantilla.Cell(fila + l, columna + c);

                            //Se comprueba si el dato es una formula
                            bool esFormula = false;
                            if(contenidoCelda is string contenidoCeldaStr && contenidoCeldaStr.StartsWith("#F#")) //Verificamos si el contenidoCelda es un string y se trata de una formula
                            {
                                esFormula = true;
                                contenidoCelda = contenidoCeldaStr.Substring(3);//Dejamos la formula sin la cadena de identificacion para poder tratarla
                            }

                            //Si el contenido es una formula, asigna el valor de la celda como una formula
                            if(esFormula)
                            {
                                cell.SetFormulaA1(contenidoCelda.ToString());
                            }
                            else
                            {
                                //Si no es una formula, se asigna el valor según su tipo original.
                                if(!string.IsNullOrEmpty(contenidoCelda.ToString()))
                                {
                                    if(contenidoCelda is int) //Entero
                                    {
                                        cell.Value = (int)contenidoCelda;
                                    }
                                    else if(contenidoCelda is decimal) //Decimal
                                    {
                                        cell.Value = Math.Round((decimal)contenidoCelda, 2); //Se redondea a 2 decimales porque en la conversion de string a float se crean muchos decimales
                                        cell.Style.NumberFormat.Format = "#,##0.00";//Se aplica el formato con 2 decimales
                                    }
                                    else if(contenidoCelda is DateTime) //Fecha
                                    {
                                        cell.Value = (DateTime)contenidoCelda;
                                        // Aplica formato personalizado para mostrar solo la fecha
                                        cell.Style.NumberFormat.Format = "dd.mm.yyyy";
                                    }
                                    else
                                    {
                                        cell.Value = contenidoCelda as string; //Resto de tipos
                                    }
                                }
                            }
                        }
                    }

                    hojaPlantilla.RecalculateAllFormulas(); //Fuerza a recalcular las formulas
                }

                catch(Exception ex)
                {
                    Program.textoLog.AppendLine("No se ha podido transformar el fichero a Excel. Revisar formulas o simbolos extraños" + ex.Message);
                }

                //Grabacion del fichero de salida
                if(File.Exists(ficheroExcel)) //Si ya existe el fichero se comprueba si hay que añadir o no hojas nuevas
                {
                    try
                    {
                        //Creacion de un nuevo fichero para grabar la salida
                        using(var ficheroSalida = new XLWorkbook(ficheroExcel))
                        {
                            IXLWorksheet hojaFicheroSalida;
                            int hojasFicheroSalida = ficheroSalida.Worksheets.Count; //Total hojas del fichero de salida
                            int hojaNueva = hojasFicheroSalida + 1; //En caso de ser necesario, numero de hoja que insertar
                            string nombreHojaSalida = $"{ficheroSalida.Worksheet(1).Name} {hojaNueva}";//Nombre de la hoja por defecto (el de la primera hoja
                            if(!Program.insertarHojas) 
                            {
                                //Si no se pasa el parametro para insertar hojas nuevas, se comprueba si la hoja pasada esta dentro del rango de hojas del fichero de salida
                                if(hoja >= 1 && hoja <= hojasFicheroSalida)
                                {
                                    //Si la hoja pasada esta dentro del rango de hojas del fichero de salida, carga el nombre que tiene la hoja pasada
                                    nombreHojaSalida = $"{ficheroSalida.Worksheet(hoja).Name}";
                                }
                                //Nota: si la hoja esta fuera del rango, hay que insertarla como nueva pero ya se ha asignado el nombre por defecto como una hoja nueva (no hace falta el else)
                            }

                            // Comprueba si la hoja ya existe
                            var hojaExistente = ficheroSalida.Worksheets.FirstOrDefault(ws => ws.Name.Equals(nombreHojaSalida, StringComparison.OrdinalIgnoreCase));

                            //Almacena la posicion de la hoja si existe, y si no la posicion sera la ultima hoja + 1
                            int posicionHoja = hojaExistente != null ? hojaExistente.Position : ficheroSalida.Worksheets.Count + 1;

                            // Se elimina la hoja si existe para evitar una excepcion al copiar los datos
                            hojaExistente?.Delete(); //El simbolo ? evita un error si la hojaExistente es null

                            // Copiar la hoja en la misma posición
                            hojaFicheroSalida = hojaPlantilla.CopyTo(ficheroSalida, nombreHojaSalida);
                            
                            //Mueve la hoja a la posicion en la que estaba antes de borrarla
                            hojaFicheroSalida.Position = posicionHoja;

                            //Graba el fichero de salida
                            ficheroSalida.Save();
                        }
                    }
                    catch(Exception ex)
                    {
                        Program.textoLog.AppendLine("No se ha podido guardar fichero Excel. Revisar si esta abierto. " + ex.Message);
                    }
                }
                //En caso de que no exista el fichero de salida
                else
                {
                    try
                    {
                        //Graba un nuevo fichero
                        using(FileStream fileOut = new FileStream(ficheroExcel, FileMode.Create))
                        {
                            libroPlantilla.SaveAs(fileOut);
                        }
                    }

                    catch(Exception ex)
                    {
                        Program.textoLog.AppendLine("No se ha podido guardar fichero Excel. Revisar si esta abierto. " + ex.Message);
                    }
                }
            }
        }



        //Metodo que permite convertir un fichero.xls (Excel 97-2003) a fichero.xlsx (Excel 2007)
        public static MemoryStream ConvertirAXlsx(string ficheroXls)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook libro = excelApp.Workbooks.Open(ficheroXls);

            // Crear un archivo temporal en formato .xlsx
            string tempFilePath = Path.GetTempFileName() + ".xlsx";
            libro.SaveAs(tempFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook);

            // Cerrar y liberar recursos
            libro.Close(false);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(libro);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            // Leer el archivo temporal en un MemoryStream
            MemoryStream ms = new MemoryStream(File.ReadAllBytes(tempFilePath));

            // Eliminar el archivo temporal
            File.Delete(tempFilePath);

            return ms;

        }
    }
}