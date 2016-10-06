using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using HtmlAgilityPack;
using System.Windows.Forms;

namespace WindowsFormsApplication4
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "html files (*.html, *.htm)|*.html; *.htm;|All files (*.*)|*.*";
            dialog.InitialDirectory = "C:\\";
            dialog.Title = "Select a text file";
            if (dialog.ShowDialog() == DialogResult.OK)
            
             */
            {
                //string num = ((((Button)sender).Name));

                //fileNamesList.Add(dialog.FileName);
                /*
                HtmlAgilityPack.HtmlDocument document2 = new HtmlAgilityPack.HtmlDocument();
                string filename = @"C:\Users\sha38475\Downloads\jhj.html";
                document2.Load(filename);
                HtmlNode[] nodes_row = document2.DocumentNode.SelectNodes("//tr").ToArray();
                HtmlNode[] nodes = document2.DocumentNode.SelectNodes("//td").ToArray();
                int cols = nodes.Length / nodes_row.Length;
                */
                HtmlAgilityPack.HtmlDocument document2 = new HtmlAgilityPack.HtmlDocument();
                string filename = @"C:\Users\sha38475\Downloads\jhj.html";
                document2.Load(filename);

                //var table = document2.DocumentNode.DescendantsAndSelf("table").Skip(1).First().Descendants("table").First();
                //var tds = table.Descendants("td");

                var table = document2.DocumentNode.SelectSingleNode("//table[1]//table[1]//table[1]//table[1]");
                HtmlNode[] node1 = document2.DocumentNode.SelectNodes("//table[1]//table[1]//table[1]//table[1]//tr").ToArray();
                label1.Text = "";
                int loop_len = node1.Length;
                for (int i = 1; i < loop_len;i++ )
                {
                    label1.Text += node1[i].InnerHtml;
                    table.AppendChild(HtmlNode.CreateNode("<tr>" + node1[i].InnerHtml + "</tr>"));
                }

                document2.Save(filename);

            }
        }

       


    }
}
