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
        static string guion = string.Empty;
        static string ficheroCsv = string.Empty;
        static string ficheroExcel = string.Empty;
        static string plantillaExcel = string.Empty;
        static int hoja = 1;
        static string celdaDestino = "A1";//Por defecto se pondra en la celda A1
        static int fila = 1;
        static int columna = 1;
        static StringBuilder textoLog = new StringBuilder();
        //static Procesos proceso = new Procesos();
        static bool agrupar = false; //Permite agrupar la salida en un solo fichero excel.

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
                    //Console.WriteLine($"Error al cargar {dllFile}: {ex.Message}");
                }
            }


            if(File.Exists("resultado.txt"))
            {
                File.Delete("resultado.txt");
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
                    List<List<object>> datos = Procesos.leerCSV(ficheroCsv); //Leer el archivo CSV
                    textoLog.AppendLine(Procesos.exportaXLSX(datos, plantillaExcel, fila, columna, hoja, ficheroExcel)); //Grabar el fichero Excel
                }
                catch(Exception ex)
                {
                    textoLog.AppendLine($"Error al procesar los ficheros: {ex.Message}");
                    grabaResultado(textoLog.ToString());
                }
            }
        }


        private static void grabaResultado(string textoLog)
        {
            //Genera un fichero con el resultado
            string ruta = Path.GetDirectoryName(ficheroExcel);
            string ficheroLog = Path.Combine(ruta, "resultado.txt");
            using(StreamWriter logger = new StreamWriter(ficheroLog))
            {
                logger.WriteLine(textoLog);
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
                    mensajeAyuda();
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
                        mensajeAyuda();
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
                        textoLog.AppendLine($"Error. No existe el fichero {guion}.");
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
                grabaResultado(textoLog.ToString());
                return false;
            }
            else
            {
                return true;
            }
        }

        static void mensajeAyuda()
        {
            StringBuilder mensaje = new StringBuilder();
            mensaje.AppendLine("Uso de la aplicacion.");
            mensaje.AppendLine();
            mensaje.AppendLine("csvTOexcel [-h] clave guion.txt");
            mensaje.AppendLine("Parametros:");
            mensaje.AppendLine("\t-h\tEsta ayuda");
            mensaje.AppendLine("\tclave\tClave de ejecucion del programa");
            mensaje.AppendLine("Contenido guion.txt");
            mensaje.AppendLine("\tENTRADA=archivo.csv (obligatorio)");
            mensaje.AppendLine("\tSALIDA=archivo.xlsx (obligatorio)");
            mensaje.AppendLine("\tPLANTILLA=plantilla.xlsx (opcional");
            mensaje.AppendLine("\tCELDA=A1 (opcional - defecto A1)");
            mensaje.AppendLine("\tHOJA=1 (opcional - defecto 1)");
            mensaje.AppendLine("\tAGRUPAR=SI (defecto NO)");
            mensaje.AppendLine();
            mensaje.AppendLine("Permite añadir formulas al CSV teniendo en cuenta lo siguiente:");
            mensaje.AppendLine("\t* El simbolo de igual se debe sustituir por #F#");
            mensaje.AppendLine("\t* La separacion de parametros de las formulas deben hacerse con comas");
            mensaje.AppendLine("\t* El nombre de las funciones debe hacerse en ingles");
            mensaje.AppendLine("\t* Ejemplo de formula (generar un hipervinculo a un fichero):");
            mensaje.AppendLine("\t #F#HYPERLINK(C:/DOCUMENTOS/000003480.PDF,000003480.PDF)");
            mensaje.AppendLine();
            mensaje.AppendLine("Pulse una tecla para salir");

            Console.Clear();
            Console.WriteLine(mensaje);
            Console.ReadKey();
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
                (clave, valor) = DivideCadena(linea, '=');

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
                                textoLog.AppendLine($"Parametros incorrectos. El fichero {ficheroCsv} no existe.");
                            }
                        }

                        break;

                    case "SALIDA":
                        ficheroExcel = valor;
                        if(string.IsNullOrEmpty(ficheroExcel))
                        {
                            textoLog.AppendLine("Parametros incorrectos. No se ha indicado el fichero de salida");
                        }
                        break;

                    case "PLANTILLA":
                        //Valor opcional (se crea un fichero Excel basico con los datos del csv)
                        plantillaExcel = valor;
                        if(!string.IsNullOrEmpty(plantillaExcel))
                        {
                            if(!File.Exists(plantillaExcel))
                            {
                                textoLog.AppendLine($"Parametros incorrectos. El fichero {plantillaExcel} no existe.");
                            }
                        }
                        break;

                    case "CELDA":
                        celdaDestino = valor.ToUpper();
                        if(!string.IsNullOrEmpty(celdaDestino))
                        {
                            int[] columnaFila = Procesos.convertirReferencia(celdaDestino);
                            fila = columnaFila[1];
                            columna = columnaFila[0];
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
                }
            }
        }

        static (string, string) DivideCadena(string cadena, char divisor)
        {
            //Permite dividir una cadena por el divisor pasado y solo la divide en un maximo de 2 partes (divide desde el primer divisor que encuentra)
            string clave = string.Empty;
            string valor = string.Empty;
            string[] partes = cadena.Split(new[] { divisor }, 2);
            if(partes.Length == 2)
            {
                clave = partes[0].Trim().ToUpper();
                valor = partes[1].Trim();
            }

            return (clave, valor);
        }
    }
}
