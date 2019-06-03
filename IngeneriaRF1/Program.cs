using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace IngeneriaRF1
{
    static class Program
    {
        private static string devRuta;
        public static string[,] Matriz = new string[50000, 3];
        private static string[] Direcciones= new string[ConfigurationManager.ConnectionStrings.Count];
        private static int cont = 0;
        private static int cantFolders = 0;
        private static int ConteoFolders = 0;
        private static int Foldersvacios = 0;
        public static ProgressBar pBar = new ProgressBar();

        //public static int cont = 0;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 



        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }


        //funcion que selecciona el archivo que contiene la ruta-------------------------------------------------------------------------------------------------------------------
        public static void Direccion()
        {
            string sFileName;
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "All Files (*.*)|*.*";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                sFileName = choofdlog.FileName;
                devRuta = File.ReadAllText(@sFileName, Encoding.UTF8);
                cantFolders = Directory.GetDirectories(devRuta).Length;
                Matriz = new string[cantFolders, 3];
                MessageBox.Show("cantidad de folders encontrados: "+cantFolders);
            }
        }



        //funcion que devuelve la ruta a utilizar por BuscarArchivos()------------------------------------------------------------------------------------------------------------   
        public static string obtener_ruta()
        { return (devRuta); }



        //buscador de los archivos------------------------------------------------------------------------------------------------------------------------------------------------
        public static void BuscarArchivos(string path, string pattern)
        {
            string[] ruta_archivo = new string[ConfigurationManager.ConnectionStrings.Count];
             
            try
            { 
                foreach (string file in Directory.GetFiles(path))
                {
                    //if (file.Contains(pattern))
                    //{
                        //Mandar el nombre de la carpeta***
                        string DirName = Nombre_carpeta(path);
                       

                        //Mandar la cantidad archivos***
                        string Cant_Archivos = cantidad_archivos(path).ToString();

                        //ruta del archivo 
                        Obtener_rutaArchivo(path);

                       
                        //puntadas del archivo***  
                        string puntadas = Obtener_Puntadas();

                        //contador que se envia para la posicion**
                        cont++;
                        
                        //Funcion para llenar Matriz 3Xinfinito***
                        llenarMatriz(DirName, Cant_Archivos, puntadas, cont);
                   
                        break;  
                    //}
                    //else
                    //{ Foldersvacios++; }
                    
                }

                foreach (string directory in Directory.GetDirectories(path) )
                {
                    BuscarArchivos(directory, pattern);
                    ConteoFolders++;                    
                   
                    if (ConteoFolders == cantFolders)
                    {                       
                        try
                        {
                            Crea_Excel();
                            MessageBox.Show("cantidad de Folders vacios es: "+Foldersvacios);
                            MessageBox.Show("Archivo De Excel Creado!!!");
                            //break;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error en la creacion de EXCELL: " + ex.Message);
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error en la busqueda de Archivos:  " + ex.Message);
            }
        }


        //funcion que llena la matriz a mostrar------------------------------------------------------------------------------------------------------------------------------------------------
        public static void llenarMatriz(string NomCap, string CantArch, string CantPun, int cont)
        {
            try
            {
                for (int f = 0; f < cont; f++)
                {
                    if (f == (cont - 1))
                    {
                        Matriz[f, 0] = NomCap;
                        Matriz[f, 1] = CantArch;
                        Matriz[f, 2] = CantPun;                        
                    }
                }
            }
            catch
            {
                Console.WriteLine("error en la entrada de la matriz!!!:");
            }
        }



        //funcion que obtiene el nombre de la carpeta que contiene archivos------------------------------------------------------------------------------------------------------------------- 
        public static string Nombre_carpeta(string path)
        {
            string FolderName = new DirectoryInfo(path).Name;
            return (FolderName);
        }



        //funcion que obtiene la cantidad de archivos en la carpeta que esta actualmente-------------------------------------------------------------------------------------------------------------------
        public static int cantidad_archivos(string path)
        {
            int count=0;
                count=Directory.GetFiles(path, "*", SearchOption.TopDirectoryOnly).Length;

            if (count==0)
            { Foldersvacios++; }

            return (count);
        }



        //funcion que obtiene la cantidad de puntadas en los archivos------------------------------------------------------------------------------------------------------------------------------------------------
        public static string Obtener_Puntadas()
        {
            string lines = "";
            int acumulador = 0;
            int count = 0;
            try
            {
                for (int D = 0; D <= Direcciones.Length; D++)
                {
                    using (StreamReader lector = new StreamReader(Direcciones[D]))
                    {                        
                        while (lector.Peek() > -1)
                        {
                            string linea = lector.ReadLine();
                            if (!String.IsNullOrEmpty(linea))
                            {
                                count++;

                                if (count == 2)
                                {
                                    lines = linea.Substring(6); //le separamos el numero

                                    acumulador += Convert.ToInt32(lines); /*lo sumamos por la cantidad de archivos*/
                                    count = 0;
                                    break;
                                }
                            }
                        }
                    }
                }                
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            //Limpiar las direcciones
            Direcciones = new string[1000];

            return (acumulador.ToString());
        }



        //crear un archivo de excell------------------------------------------------------------------------------------------------------------------------------------------------
        public static void Crea_Excel()
        {            
            try
            {
                using (ExcelPackage excel = new ExcelPackage())
                {
                    //Crear las 3 instancias de EXCELL----------------------------------------------------------
                    excel.Workbook.Worksheets.Add("Worksheet1");
                    excel.Workbook.Worksheets.Add("Worksheet2");
                    excel.Workbook.Worksheets.Add("Worksheet3");

                    //Agregar header al EXCELL----------------------------------------------------------
                    var headerRow = new List<string[]>()
                    {
                    new string[] { "Nombre de la Carpeta", "Cantidad de Archivos", "Puntadas"}
                    }; 

                    // Determine the header range (e.g. A1:C1)
                    string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                    // Target a worksheet
                    var worksheet = excel.Workbook.Worksheets["Worksheet1"];                   

                    // Popular header row data
                    worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                    //Agregar la matriz al EXCELL-------------------------------------------------------                  
                    int contador = 1;

                    //instancia donde se crea el objeto a ingresar al excell
                    var cellData= new List<string[]>();
                    for (int f = 0; f < Matriz.GetLength(0); f++)
                    {

                        contador++;
                        string posicion = "A" + contador.ToString() + ":";

                             cellData = new List<string[]>()
                            {
                              new string[] { Matriz[f,0], Matriz[f,1], Matriz[f,2] }
                            };                      

                        // Determine the header range(e.g.A2:C2)
                        string MatrisRange = posicion + Char.ConvertFromUtf32(cellData[0].Length + 64) + contador;

                        // Target a worksheet
                        var worksheet2 = excel.Workbook.Worksheets["Worksheet1"];

                        // Popular header row data
                        worksheet.Cells[MatrisRange].LoadFromArrays(cellData);
                    }


                    // FileInfo excelFile = new FileInfo(@"C:\Users\greva\Desktop\Proyecto de IngeneriaRF1\puebas\PuntadasByOrden.xlsx");
                    verificar_carpeta();
                    string thisDay = DateTime.Now.ToString("hh-mm-ss");
                    FileInfo excelFile = new FileInfo(@"C:\RecordPuntadas\PuntadasByOrden-("+ thisDay + ").xlsx");
                    excel.SaveAs(excelFile);
                    limpiar_variables();
                }
            }
            catch (Exception ex)
            { MessageBox.Show("Error en la creacion del EXCELL:" + ex.Message); }

        }

        public static void Obtener_rutaArchivo(string ruta_general)
        {

            DirectoryInfo di = new DirectoryInfo(ruta_general);
            int count = 0;

            foreach (var fi in di.GetFiles("*.dst"))
            {
                Direcciones[count] = @"" + ruta_general.ToString() + "\\" + fi.Name.ToString();              
                count++;
            }
        }

        private static void verificar_carpeta()
        {

            string rutaFolder =@"C:\RecordPuntadas";

            try
            {
                //si no existe la carpeta temporal la creamos 
                if (!(Directory.Exists(rutaFolder)))/*carpeta*/
                {
                    Directory.CreateDirectory(rutaFolder);
                    MessageBox.Show("carpeta creada");
                }

            }
            catch (Exception errorC)
            {
                MessageBox.Show("Ha habido un error al intentar " +
                         "crear el fichero temporal:" +
                         Environment.NewLine + Environment.NewLine +/*
                         rutafolder +*/ Environment.NewLine +
                         Environment.NewLine + errorC.Message,
                         "Error al crear fichero temporal",
                         MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private static void limpiar_variables()
        {
             devRuta="";
         Matriz = new string[50000, 3];
         Direcciones = new string[ConfigurationManager.ConnectionStrings.Count];
        cont = 0;
         cantFolders = 0;
        ConteoFolders = 0;
        //public static ProgressBar pBar = new ProgressBar();
    }
        //aqui llega la ultima funcion...
    }
}
