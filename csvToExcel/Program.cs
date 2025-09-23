using System.Collections.Generic;
using System.IO;
using System;
using System.Reflection;
using System.Text;

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
        public static string guion = string.Empty;
        public static string ficheroCsv = string.Empty;
        public static string ficheroExcel = string.Empty;
        public static string plantillaExcel = string.Empty;
        public static int hoja = 1;
        public static string celdaDestino = "A1";//Por defecto se pondra en la celda A1
        public static int fila = 1;
        public static int columna = 1;
        public static StringBuilder textoLog = new StringBuilder(); //Texto que almacena los posibles errores
        public static bool agrupar = false; //Permite agrupar la salida en un solo fichero excel.
        public static bool insertarHojas = true; //Controla si hay que insertar o no hojas nuevas

        static void Main(string[] args)
        {
            //Ruta donde estan ubicadas las librerias necesarias
            string libsDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "dse_dlls");

            // Cargar las librerías dinámicamente
            foreach(var dllFile in Directory.GetFiles(libsDirectory, "*.dll"))
            {
                try
                {
                    Assembly.LoadFrom(dllFile); // Carga las librerías
                }
                catch(Exception ex)
                {
                    textoLog.AppendLine($"Error al cargar la libreria {dllFile}: {ex.Message}");
                }
            }

            //Borra el fichero de errores si existe de una ejecucion anterior
            if(File.Exists("errores.txt"))
            {
                File.Delete("errores.txt");
            }

            //Procesa el guion y carga los datos
            if(gestionParametros(args))
            {
                try
                {
                    if(!agrupar) //Si no se agrupa en un solo fichero, se elimina el fichero de salida para evitar que se añadan hojas
                    {
                        if(File.Exists(ficheroExcel)) File.Delete(ficheroExcel);
                    }

                    //Lee el archivo CSV
                    List<List<object>> datos = Procesos.leerCSV(ficheroCsv);

                    //Proceso para grabar el fichero excel de salida
                    Procesos.exportaXLSX(datos);
                }
                catch(Exception ex)
                {
                    textoLog.AppendLine($"Error al procesar los ficheros: {ex.Message}");
                }
            }

            if(textoLog.Length > 0)
            {
                Utilidades.grabaResultado();
            }
        }

        static bool gestionParametros(string[] parametros)
        {
            int totalParametros = parametros.Length;
            switch(totalParametros)
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
                    Utilidades.mensajeAyuda();
                    Console.SetWindowSize(120, 28);
                    Console.SetBufferSize(120, 28);
                    Console.ResetColor();
                    Console.Clear();
                    Environment.Exit(0);
                    break;

                case 1:
                    //Si solo se pasa un parametro puede ser la peticion de ayuda
                    if(parametros[0] == "-h")
                    {
                        Utilidades.mensajeAyuda();
                        return false;
                    }
                    break;

                case 2:
                    //Si se pasan dos parametros el primero debe ser la clave de ejecucion y el segundo el fichero con el guion
                    if(parametros[0] != "ds123456")
                    {
                        textoLog.AppendLine("Clave de ejecucion del programa incorrecta");
                        return false;
                    }

                    guion = parametros[1];
                    if(!File.Exists(guion))
                    {
                        textoLog.AppendLine($"Parametros incorrectos. No existe el fichero del guion {guion}");
                        return false;
                    }
                    break;

                default:
                    textoLog.AppendLine("Parametros incorrectos. Consulte la ayuda");
                    return false;

            }

            CargaGuion();

            if(textoLog.Length > 0)
            {
                //Se han producido errores en el procesado
                return false;
            }
            else
            {
                return true;
            }
        }

        //Metodo para procesar y almacenar el guion
        static void CargaGuion()
        {
            //Procesa las lineas del guion
            foreach(string linea in File.ReadAllLines(guion))
            {
                //Evita procesar lineas vacias
                if(string.IsNullOrWhiteSpace(linea)) continue;

                //Divide la linea en clave=valor
                string clave = string.Empty;
                string valor = string.Empty;
                (clave, valor) = Utilidades.DivideCadena(linea, '=');

                switch(clave)
                {
                    case "ENTRADA":
                        ficheroCsv = valor;
                        if(string.IsNullOrEmpty(ficheroCsv))
                        {
                            textoLog.AppendLine("Parametros incorrectos. No se ha indicado el fichero de entrada");
                        }
                        else
                        {
                            if(!File.Exists(ficheroCsv))
                            {
                                textoLog.AppendLine($"Parametros incorrectos. No existe el fichero de entrada {ficheroCsv}");
                            }
                        }

                        break;

                    case "SALIDA":
                        ficheroExcel = valor;
                        if(string.IsNullOrEmpty(ficheroExcel))
                        {
                            textoLog.AppendLine("Parametros incorrectos. No se ha indicado el fichero de salida");
                        }
                        ficheroExcel = Path.ChangeExtension(ficheroExcel, "xlsx");
                        break;

                    case "PLANTILLA":
                        //Valor opcional (se crea un fichero Excel basico con los datos del csv)
                        plantillaExcel = valor;
                        if(!string.IsNullOrEmpty(plantillaExcel))
                        {
                            if(!File.Exists(plantillaExcel))
                            {
                                //Controla que exista el fichero de la plantilla pasada como parametro
                                textoLog.AppendLine($"Parametros incorrectos. No existe el fichero con la plantilla {plantillaExcel}");
                            }
                        }
                        break;

                    case "CELDA":
                        celdaDestino = valor.ToUpper();
                        if(!string.IsNullOrEmpty(celdaDestino))
                        {
                            int[] columnaFila = Utilidades.convertirReferencia(celdaDestino);
                            columna = columnaFila[0];
                            fila = columnaFila[1];
                        }
                        break;

                    case "HOJA":
                        hoja = Convert.ToInt32(valor);
                        if(hoja < 1)
                        {
                            textoLog.AppendLine("El numero de hoja no puede ser menor de 1");
                        }
                        break;

                    case "AGRUPAR":
                        string opcion = valor.ToUpper();
                        if(opcion == "SI")
                        {
                            agrupar = true;
                        }
                        break;

                    case "INSERTAR":
                        string insertar = valor.ToUpper();
                        if(insertar == "NO")
                        {
                            insertarHojas = false;
                        }
                        break;
                }
            }
        }
    }
}
