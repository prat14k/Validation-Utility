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
            Button b = (Button)sender;
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
            int numberOfFiles = fileNamesList.Count();
            
            if (HeaderFile == "")
                msg = "Nothing To do !!!";
            else if (numberOfFiles == 0) {
                msg = "File converted to Excel and stored at=> " + convert2Excel(HeaderFile);
            }
            else
            {
                HtmlAgilityPack.HtmlDocument document1 = new HtmlAgilityPack.HtmlDocument();
                HtmlAgilityPack.HtmlDocument document2 = new HtmlAgilityPack.HtmlDocument();
                
                string tempFilePath = HeaderFile;
             
                document1.Load(tempFilePath);
                var table = document1.DocumentNode.SelectSingleNode("//table[1]//table[1]//table[1]//table[1]");
                        
                for (int l = 0; l < numberOfFiles; l++)
                {
                    {
                        document2.Load(fileNamesList[l]);
                        HtmlNode table2 = document2.DocumentNode.SelectSingleNode("//table[1]//table[1]//table[1]//table[1]");
                        HtmlNode td = document2.DocumentNode.SelectSingleNode("//table[1]//table[1]//table[1]//table[1]//td[1]");
                        HtmlNode th = document2.DocumentNode.SelectSingleNode("//table[1]//table[1]//table[1]//table[1]//tr[1]");
                        //for (int j = 1; j < node1.Length; j++)
                            table2.RemoveChild(td);
                            table2.RemoveChild(th);
                            table.AppendChild(HtmlNode.CreateNode(table2.InnerHtml.Trim()));
                            //counter++;
                    }
                    
                }

                tempFilePath = Path.GetDirectoryName(HeaderFile) + "\\Temp_" + Path.GetFileName(HeaderFile);
                document1.Save(tempFilePath);
                convert2Excel(tempFilePath);
                File.Delete(tempFilePath);
            }
            MessageBox.Show(msg);
            
        }

        private string convert2Excel(string tempFilePath)
        {
            string finalPath="";
            var app = new Excel.Application();
            try
            {
                var wb = app.Workbooks.Open(tempFilePath);
                
                if (Path.GetExtension(HeaderFile) == ".html")
                    finalPath = HeaderFile.Replace(".html", ".xlsx");
                else
                    finalPath = HeaderFile.Replace(".htm", ".xlsx");
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

        
        
    }
}
