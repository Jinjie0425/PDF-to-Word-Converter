using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.WindowsAPICodePack.Dialogs;

//support pdf to word process
using Spire.Pdf;
using System.IO;
using System.Diagnostics;

namespace Pdf_to_Word
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        string path = "";
        string destination = "";

        //File Browse Button
        private void browse_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            //fd.Filter = "pdf files (*.pdf) | *.pdf | All files (*.*) | (*.*);";
            fd.Filter = "pdf files (*.pdf) | *.pdf;";
            if (fd.ShowDialog() == DialogResult.OK)
            {
                path = fd.FileName;
                //Text inside box equals to filename
                textBox1.Text = fd.FileName;
            }
        }

        //Convert Button
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                PdfDocument obj = new PdfDocument();
                //Load the selected file
                obj.LoadFromFile(path);
                obj.SaveToFile("Convert.docx", FileFormat.DOCX);

                /*//Select location function
                try
                {
                    if (File.Exists("Convert.docx") == true)
                    {
                        //Open document after converted
                        //Process.Start("Convert.docx");
                        string file = @"Convert.docx;";
                        //Your program location
                        string fileToMove = @"C:\Users\User\source\repos\Pdf_to_Word\Pdf_to_Word\bin\Debug\netcoreapp3.1\" + "Convert.docx";
                        string moveTo = destination;
                        //moving file
                        File.Move(fileToMove, moveTo);
                    }
                }

                catch (Exception ex)
                {
                    System.Console.WriteLine(ex.Message);
                }*/
            }

            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);
            }
        }

        //Destination Browser Button
        private void button1_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\Users";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                destination = dialog.FileName;
                //Text inside box equals to filename
                textBox2.Text = dialog.FileName;
            }

        }
    }
}
