using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HtmlAgilityPack;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace WindowsFormsApplication4
{
    public partial class Form1 : Form
    {
        bool typeOfOutput = false;
        public static string msg ="Merged !!!!";
        public static string HeaderFile = "";
        public static List<string> fileNamesList = new List<string>();
        //public static string[] fileNames = new string[60];

        public Form1()
        {
            InitializeComponent();
        }


        private void button_Click(object sender, EventArgs e)
        {
            {

                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Multiselect = true;
                dialog.Filter = "html files (*.html, *.htm)|*.html; *.htm;|All files (*.*)|*.*";
                
                dialog.Title = "Select the Validation file";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    fileNamesList.Clear();
                    listBox1.Items.Clear();
                    
                    HeaderFile = dialog.FileNames[0];
                    listBox1.Items.Add(HeaderFile);
                    for (int i = 1; i < dialog.FileNames.Length;i++ )
                    {
                        fileNamesList.Add(dialog.FileNames[i]);
                        listBox1.Items.Add(dialog.FileNames[i]);
                    }
                    
                }
            }
        }

        private void mergeButton_Click(object sender, EventArgs e)
        {
            Button b = (Button)sender;
            
            int numberOfFiles = fileNamesList.Count();

            if (HeaderFile == "")
            {
                msg = "Nothing To do !!!";
                MessageBox.Show(msg);
            }
            else if (numberOfFiles == 0)
            {
                msg = "File converted to Excel and stored at=> " + convert2Excel(HeaderFile);
                MessageBox.Show(msg);
            }

            else
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                if (b.Name == "html")
                {
                    saveFileDialog1.Filter = "HTML File|*.html; *.htm|All files (*.*)|*.*";
                    saveFileDialog1.Title = "Save an HTML File";
                }
                else if (b.Name == "excelToo")
                {
                    saveFileDialog1.Filter = "All files (*.*)|*.*";
                    saveFileDialog1.Title = "Save an HTML as well as Excel File";

                }

                // If the file name is not an empty string open it for saving.
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (saveFileDialog1.FileName != "")
                    {
                        string temp = Path.GetDirectoryName(saveFileDialog1.FileName) + "\\" + Path.GetFileNameWithoutExtension(saveFileDialog1.FileName) + ".html";
                        string[] lines = File.ReadAllLines(HeaderFile);
                        int totalLines = lines.Length;
                        string[] appendLines = new string[8];
                        for (int i = 1; i <= 8; i++)
                        {
                            appendLines[8 - i] = lines[totalLines - i];
                            lines[totalLines - i] = "";
                        }

                        File.WriteAllLines(temp, lines);
                        string[] arr;
                        for (int l = 0; l < numberOfFiles; l++)
                        {
                            {
                                arr = File.ReadAllLines(fileNamesList[l]);
                                totalLines = arr.Length;
                                for (int i = 0; i < 38; i++)
                                    arr[i] = "";
                                for (int i = 1; i <= 8; i++)
                                {
                                    arr[totalLines - i] = "";
                                }
                                File.AppendAllLines(temp, arr);
                            }

                        }
                        File.AppendAllLines(temp, appendLines);
                        msg = "File is merged and stored at => " + temp;
                        if (b.Name == "excelToo")
                            convert2Excel(temp);
                    }
                    else
                        msg = "File name for saving was empty !!!";
                    MessageBox.Show(msg);
                }

            }
                        
        }

        private string convert2Excel(string tempFilePath)
        {
            string finalPath="";
            var app = new Excel.Application();
            try
            {
                var wb = app.Workbooks.Open(tempFilePath);
                
                finalPath = tempFilePath.Replace(".html", ".xlsx");
                
                wb.SaveAs(finalPath, Excel.XlFileFormat.xlOpenXMLWorkbook);

                wb.Close();
                msg = "File is merged and stored at => " + finalPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                app.Quit();
                
            }
            return finalPath;
        }

        private void label1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This is a utility to merge different html files and output html or excel. If only one file selected,it will output only an excel file. But if more are selected, it is upto you if you want .html only or both .html or .xlsx !!!!!");
        }

    }
}
