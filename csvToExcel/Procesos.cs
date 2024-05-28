using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Ude;


namespace csvToExcel
{

    public class Procesos
    {
        //Metodo para leer el CSV creando una lista de objetos que almacenara las lineas y dentro otra lista de objetos con cada uno de los campos
        public List<List<object>> leerCSV(string archivoCSV)
        {
            string[] lines;

            try
            {
                //Deteccion de la codificacion del fichero csv
                var encodingDetector = new CharsetDetector();
                using (var filestream = new FileStream(archivoCSV, FileMode.Open))
                {
                    encodingDetector.Feed(filestream);
                    encodingDetector.DataEnd();

                }

                //Almacena la codificacion en la variable 'codificacion' para luego pasarla como parametro al leer el fichero y evitar que los caracteres como la 'Ñ' salgan con simbolos.
                var charset = encodingDetector.Charset;
                var encoding = charset != null ? Encoding.GetEncoding(charset) : Encoding.Default;

                //Almacena todas las lineas del archivoCSV en la variable de array 'lines' 
                using (StreamReader fichero = new StreamReader(archivoCSV, encoding))
                {
                    lines = fichero.ReadToEnd().Split('\n');
                }

                //Almacena todos los campos de la linea en la variable de array 'datos' 
                List<List<object>> datos = new List<List<object>>();

                for (int i = 0; i < lines.Length; i++)
                {
                    if (!string.IsNullOrEmpty(lines[i])) //Evita almacenar lineas vacias
                    {
                        //Divide cada linea en campos con el separador ';'. Ademas se establece cada campo con el tipo de valor que le corresponde segun su valor (int, float, DateTime, etc.
                        List<object> linea = lines[i].Split(';').Select(x => tipoValor(x)).ToList<object>();
                        datos.Add(linea); //Se inserta en cada linea todos los campos en los que se haya dividido
                    }
                }

                return datos;
            }

            catch (Exception ex)
            {
                throw new Exception("Error al procesar el CSV" + ex);
            }
        }

        //Convierte cada objeto 'string' al tipo que le corresponde segun su valor.
        private object tipoValor(string value)
        {
            if (int.TryParse(value, out int intValue))
            {
                return intValue;
            }
            else if (float.TryParse(value, out float floatValue))
            {
                return floatValue;
            }
            else if (DateTime.TryParse(value, out DateTime dateTimeValue))
            {
                return dateTimeValue;
            }
            else
            {
                return value; // Si no se puede convertir, se mantiene como string
            }
        }


        public string exportaXLSX(List<List<object>> datos, string plantillaExcel, int fila, int columna, int hoja, string ficheroExcel)
        {
            string resultado = string.Empty;
            string nombreFichero = string.IsNullOrEmpty(plantillaExcel) ? "fichero.xlsx" : plantillaExcel;

            //Creacion del libro y hoja
            using (FileStream file = new FileStream(nombreFichero, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                XLWorkbook libroPlantilla;
                IXLWorksheet hojaPlantilla;
                try
                {
                    //Cargar la plantilla y si no existe crear un libro nuevo vacio
                    if (string.IsNullOrEmpty(plantillaExcel))
                    {
                        libroPlantilla = new XLWorkbook(); //Si no se pasa la plantilla se crea un libro nuevo
                        hojaPlantilla = libroPlantilla.Worksheet(1); //Se crea la hoja 1 ya que no existe el libro
                    }
                    else
                    {
                        {
                            libroPlantilla = new XLWorkbook(file);
                            hojaPlantilla = libroPlantilla.Worksheet(hoja);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new IndexOutOfRangeException(ex.Message + $" Fichero = {ficheroExcel}. Hoja = {hoja}");
                }

                try
                {
                    // Escribir los datos en la plantilla
                    for (int l = 0; l < datos.Count; l++)
                    {
                        for (int c = 0; c < datos[l].Count; c++)
                        {
                            object contenidoCelda = datos[l][c];
                            var cell = hojaPlantilla.Cell(fila + l, columna + c);
                            //Se comprueba si el dato es una formula
                            bool esFormula = false;
                            if (contenidoCelda is string contenidoCeldaStr && contenidoCeldaStr.StartsWith("#F#")) //Verificamos si el contenidoCelda es un string y se trata de una formula
                            {
                                esFormula = true;
                                contenidoCelda = contenidoCeldaStr.Substring(3);//Dejamos la formula sin la cadena de identificacion para poder tratarla
                            }

                            if (esFormula)
                            {
                                cell.SetFormulaA1(contenidoCelda.ToString());//Grabamos la formula con el contenido del objeto
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(contenidoCelda.ToString()))
                                {
                                    // Si no es una fórmula, se asigna el valor según su tipo original.
                                    if (contenidoCelda is int) //Entero
                                    {
                                        cell.Value = (int)contenidoCelda;
                                    }
                                    else if (contenidoCelda is float) //Decimal
                                    {
                                        cell.Value = Math.Round((float)contenidoCelda, 2); //Se redondea a 2 decimales porque en la conversion de string a float se crean muchos decimales
                                        cell.Style.NumberFormat.Format = "#,##0.00";//Se aplica el formato con 2 decimales
                                    }
                                    else if (contenidoCelda is DateTime) //Fecha
                                    {
                                        cell.Value = (DateTime)contenidoCelda;
                                        // Aplicar formato personalizado para mostrar solo la fecha
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

                catch (Exception ex)
                {
                    throw new Exception("No se ha podido transformar el fichero a Excel. Revisar formulas o simbolos extraños" + ex.Message);
                }

                //Grabacion del fichero de salida
                if (File.Exists(ficheroExcel)) //Si ya existe el fichero en la rura se añaden hojas con el nuevo procesado
                {
                    try
                    {
                        using (var ficheroSalida = new XLWorkbook(ficheroExcel))
                        {
                            int hojaNueva = ficheroSalida.Worksheets.Count + 1;
                            string nombreHojaNueva = $"Informe Planning {hojaNueva}";
                            var hojaFicheroSalida = hojaPlantilla.CopyTo(ficheroSalida, nombreHojaNueva);
                            ficheroSalida.Save();
                            resultado = $"OK. Fichero '{ficheroExcel}' generado, y añadida la hoja {nombreHojaNueva}";
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("No se ha podido guardar fichero Excel. Revisar si esta abierto. " + ex.Message);
                    }
                }

                else
                {
                    try
                    {
                        using (FileStream fileOut = new FileStream(ficheroExcel, FileMode.Create))
                        {
                            libroPlantilla.SaveAs(fileOut);
                            resultado = $"OK. Fichero '{ficheroExcel}' generado";

                        }
                    }

                    catch (Exception ex)
                    {
                        throw new Exception("No se ha podido guardar fichero Excel. Revisar si esta abierto. " + ex.Message);
                    }
                }
            }

            return resultado;
        }

        public int[] convertirReferencia(string celdaRef)
        {
            int[] referencia = new int[2];

            string colRef = string.Empty;
            string rowRef = string.Empty;

            foreach (char c in celdaRef)
            {
                if (char.IsLetter(c))
                {
                    colRef += c;
                }
                else if (char.IsDigit(c))
                {
                    rowRef += c;
                }
            }

            int fila = int.Parse(rowRef);
            int columna = 0;
            foreach (char c in colRef)
            {
                columna = columna * 26 + (c - 'A' + 1);
            }

            referencia[0] = columna;
            referencia[1] = fila;

            return referencia;
        }
    }
}