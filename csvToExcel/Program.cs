using System.Collections.Generic;
using System.IO;
using System;
using System.Reflection;

namespace csvToExcel
{
    class Program
    {
        // Obtener la información del ensamblado actual
        static Assembly assembly = Assembly.GetExecutingAssembly();

        // Obtener el atributo del copyright del ensamblado
        static AssemblyCopyrightAttribute copyrightAttribute =
            (AssemblyCopyrightAttribute)Attribute.GetCustomAttribute(assembly, typeof(AssemblyCopyrightAttribute));

        //Obtener el atributo del nombre del ensamblado
        static AssemblyProductAttribute nombreProducto = (AssemblyProductAttribute)Attribute.GetCustomAttribute(assembly, typeof(AssemblyProductAttribute));

        // Obtener el valor de la propiedad Copyright y nombre del producto
        static string copyrightValue = copyrightAttribute?.Copyright;
        static string nombreValue = nombreProducto?.Product;

        //Variable para chequear si los parametros pasados son correctos
        static bool continuar = false;

        //Variables para gestion de ficheros
        static string ficheroCsv = string.Empty;
        static string ficheroExcel = string.Empty;
        static string plantillaExcel = string.Empty;
        static int hoja = 1;
        static string celdaDestino = "A1";//Por defecto se pondra en la celda A1
        static int fila = 1;
        static int columna = 1;
        static string textoLog = string.Empty;
        static Procesos proceso = new Procesos();

        static void Main(string[] args)
        {
            if (File.Exists("resultado.txt"))
            {
                File.Delete("resultado.txt");
            }
            continuar = gestionParametros(args);
            if (continuar)
            {
                try
                {
                    List<List<object>> datos = proceso.leerCSV(ficheroCsv); //Leer el archivo CSV
                    textoLog += proceso.exportaXLSX(datos, plantillaExcel, fila, columna, hoja, ficheroExcel); //Grabar el fichero Excel
                }
                catch (Exception ex)
                {
                    textoLog += "Error al procesar los ficheros: " + ex.Message + "\n";
                    //Console.WriteLine("Error: " + ex.Message);
                    grabaResultado(textoLog);
                }
            }
        }


        private static void grabaResultado(string textoLog)
        {
            //Genera un fichero con el resultado
            string ficheroLog = "resultado.txt";
            using (StreamWriter logger = new StreamWriter(ficheroLog))
            {
                logger.WriteLine(textoLog);
            }
        }

        static bool gestionParametros(string[] parametros)
        {
            int totalParametros = parametros.Length;
            int controlParametros = 0;
            switch (totalParametros)
            {
                case 0:
                    //Si no se pasan argumentos debe ser porque se ha ejecutado desde windows
                    // Abre una ventana de consola para mostrar el mensaje
                    Console.BackgroundColor = ConsoleColor.DarkRed;
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.SetWindowSize(120, 28);
                    Console.SetBufferSize(120, 28);
                    Console.Clear();
                    Console.Title = $"{nombreValue} - {copyrightValue}";
                    Console.WriteLine("\r\nEsta aplicacion debe ejecutarse por linea de comandos y pasarle los parametros correspondientes.");
                    mensajeAyuda();
                    Console.SetWindowSize(120, 28);
                    Console.SetBufferSize(120, 28);
                    Console.ResetColor();
                    Console.Clear();
                    Environment.Exit(0);
                    break;

                case 1:
                    //Si solo se pasa un parametro puede ser la peticion de ayuda
                    if (parametros[0] == "-h")
                    {
                        mensajeAyuda();
                    }
                    break;

                default:
                    break;
            }


            foreach (string parametro in parametros)
            {
                if (parametro == "-h")
                {
                    mensajeAyuda();
                    return false;
                }
                string[] partes = parametro.Split('=');
                if (partes.Length == 2)
                {
                    string key = partes[0].ToLower();
                    string value = partes[1];

                    switch (key)
                    {
                        case "entrada":
                            ficheroCsv = value;
                            if (string.IsNullOrEmpty(ficheroCsv))
                            {
                                textoLog += "Parametros incorrectos. No se ha indicado el fichero de entrada\r\n";
                            }
                            else
                            {
                                if (!File.Exists(ficheroCsv))
                                {
                                    textoLog += $"Parametros incorrectos. El fichero {ficheroCsv} no existe.\r\n";
                                }
                                else
                                {
                                    controlParametros++;
                                }
                            }
                            break;

                        case "salida":
                            ficheroExcel = value;
                            if (string.IsNullOrEmpty(ficheroExcel))
                            {
                                textoLog += "Parametros incorrectos. No se ha indicado el fichero de salida\r\n";
                            }
                            else
                            {
                                controlParametros++;
                            }
                            break;

                        case "plantilla":
                            plantillaExcel = value;
                            break;

                        case "celda":
                            celdaDestino = value.ToUpper();
                            if (!string.IsNullOrEmpty(celdaDestino))
                            {
                                int[] columnaFila = proceso.convertirReferencia(celdaDestino);
                                fila = columnaFila[1];
                                columna = columnaFila[0];
                            }
                            break;

                        case "hoja":
                            hoja = Convert.ToInt32(value);
                            if (hoja < 1)
                            {
                                textoLog += "El numero de hoja no puede ser menor de 1\r\n";
                            }
                            break;
                    }
                }
            }

            if (controlParametros == 2)
            {
                return true;
            }
            else
            {
                grabaResultado(textoLog);
                return false;
            }
        }

        static void mensajeAyuda()
        {
            string mensaje =
                "\nUso de la aplicacion.\r\n\r\n" +
                "csvTOexcel -h [parametro1 parametro2 ... parametroN]\r\n" +
                "\r\nParametros:\r\n" +
                "\t-h\tEsta ayuda\r\n" +
                "\tentrada=archivo.csv (obligatorio)\r\n" +
                "\tsalida=archivo.xlsx (obligatorio)\r\n" +
                "\tplantilla=plantilla.xlsx (opcional\r\n" +
                "\tcelda=A1 (defecto)\r\n" +
                "\thoja=1 (defecto)\r\n\r\n" +
                "Permite añadir formulas al CSV teniendo en cuenta lo siguiente:\r\n" +
                "\t* El simbolo de igual se debe sustituir por #F# \r\n" +
                "\t* La separacion de parametros de las formulas deben hacerse con comas\r\n" +
                "\t* El nombre de las funciones debe hacerse en ingles\r\n" +
                "\t* Ejemplo de formula (generar un hipervinculo a un fichero):\r\n" +
                "\t #F#HYPERLINK(C:/DOCUMENTOS/000003480.PDF,000003480.PDF)" +
                "\r\n\r\nPulse una tecla para salir";
            Console.Clear();
            Console.WriteLine(mensaje);
            Console.ReadKey();
        }
    }
}
