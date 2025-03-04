using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace csvToExcel
{
    public static class Utilidades
    {
        //Graba el contenido el textoLog que tiene los errores del proceso
        public static void grabaResultado()
        {
            //Genera un fichero con los errores 
            string ruta = string.Empty;
            if(!string.IsNullOrEmpty(Program.ficheroExcel))
            {
                ruta = Path.GetDirectoryName(Program.ficheroExcel);
            }
            string ficheroLog = Path.Combine(ruta, "errores.txt");
            using(StreamWriter logger = new StreamWriter(ficheroLog))
            {
                logger.WriteLine(Program.textoLog);
            }
        }

        //Muestra el mensaje de ayuda
        public static void mensajeAyuda()
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
            mensaje.AppendLine("\tAGRUPAR=SI (defecto NO). Añade hojas al final del fichero o borra (valor 'NO') el fichero previamente");
            mensaje.AppendLine("\tINSERTAR=NO (defecto SI). Copia los datos en el fichero segun la hoja pasada como parametro (no añade hojas) o añade hojas nuevas al final del fichero (valor 'SI'");
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

        public static int[] convertirReferencia(string celdaRef)
        {
            int[] referencia = new int[2];

            string colRef = string.Empty;
            string rowRef = string.Empty;

            foreach(char c in celdaRef)
            {
                if(char.IsLetter(c))
                {
                    colRef += c;
                }
                else if(char.IsDigit(c))
                {
                    rowRef += c;
                }
            }

            int fila = int.Parse(rowRef);
            int columna = 0;
            foreach(char c in colRef)
            {
                columna = columna * 26 + (c - 'A' + 1);
            }

            referencia[0] = columna;
            referencia[1] = fila;

            return referencia;
        }

        public static (string, string) DivideCadena(string cadena, char divisor)
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
