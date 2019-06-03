using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

// This is the code for your desktop app.
// Press Ctrl+F5 (or go to Debug > Start Without Debugging) to run your app.

namespace IngeneriaRF1
{
    public partial class Form1 : Form
    {
        private int conteo;
        public Form1()
        {
            InitializeComponent();
            conteo = 0;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // Click on the link below to continue learning how to build a desktop app using WinForms!
            System.Diagnostics.Process.Start("http://aka.ms/dotnet-get-started-desktop");

        }

        public void button1_Click(object sender, EventArgs e)
        {
           
            //Se envian los datos necesarios para escanear los archivos 
            string extencion ="dst";
            string path = Program.obtener_ruta();
            //progressBar1.Value = 0;

            Program.BuscarArchivos(path,extencion);


        }

        private void button2_Click(object sender, EventArgs e)
        {

            //Se llaman al programa que selecionara al archivo que contiene la ruta
            Program.Direccion();
        }

        private void helloWorldLabel_Click(object sender, EventArgs e)
        {

        }

        public void timer1_Tick(object sender, EventArgs e)
        {
            //conteo++;
            //lblporciento.Text = (conteo.ToString() + " %");
            //if (progressBar1.Value < 100)
            //    progressBar1.Value++;
            //else
            //    timer1.Enabled = false;
        }
    }
}
