using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HtmlAgilityPack;

namespace WindowsFormsApplication4
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "html files (*.html, *.htm)|*.html; *.htm;|All files (*.*)|*.*";
            dialog.Title = "Select the Validation file";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string tempFilePath = dialog.FileName;
                HtmlAgilityPack.HtmlDocument document1 = new HtmlAgilityPack.HtmlDocument();
                document1.Load(tempFilePath);
                HtmlNode table = document1.DocumentNode.SelectSingleNode("//table[1]//table[1]//table[1]//table[1]");
                HtmlNode td = document1.DocumentNode.SelectSingleNode("//table[1]//table[1]//table[1]//table[1]//td[1]");
                HtmlNode th = document1.DocumentNode.SelectSingleNode("//table[1]//table[1]//table[1]//table[1]//tr[1]");
                table.RemoveChild(td);
                table.RemoveChild(th);
                label1.Text = table.InnerHtml.TrimStart();

            }
                 
        }



    }
}
