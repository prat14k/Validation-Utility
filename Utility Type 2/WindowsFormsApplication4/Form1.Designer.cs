namespace WindowsFormsApplication4
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.title = new System.Windows.Forms.Label();
            this.description = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.html = new System.Windows.Forms.Button();
            this.excelToo = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // title
            // 
            this.title.AutoSize = true;
            this.title.Font = new System.Drawing.Font("Microsoft Sans Serif", 18.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.title.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.title.Location = new System.Drawing.Point(247, 29);
            this.title.Name = "title";
            this.title.Size = new System.Drawing.Size(229, 29);
            this.title.TabIndex = 0;
            this.title.Text = "File Merger Utility";
            // 
            // description
            // 
            this.description.AutoSize = true;
            this.description.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.description.Location = new System.Drawing.Point(168, 67);
            this.description.Name = "description";
            this.description.Size = new System.Drawing.Size(390, 20);
            this.description.TabIndex = 1;
            this.description.Text = "Merge html file tables, Convert them to excel easily";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(252, 106);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(196, 35);
            this.button1.TabIndex = 3;
            this.button1.Text = "Select File to Merge";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button_Click);
            // 
            // listBox1
            // 
            this.listBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBox1.FormattingEnabled = true;
            this.listBox1.HorizontalScrollbar = true;
            this.listBox1.ItemHeight = 17;
            this.listBox1.Location = new System.Drawing.Point(61, 159);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(606, 497);
            this.listBox1.TabIndex = 13;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Fluor Daniel", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.DarkCyan;
            this.label1.Location = new System.Drawing.Point(670, 122);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 19);
            this.label1.TabIndex = 14;
            this.label1.Text = "Help";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // html
            // 
            this.html.Location = new System.Drawing.Point(190, 680);
            this.html.Name = "html";
            this.html.Size = new System.Drawing.Size(99, 50);
            this.html.TabIndex = 15;
            this.html.Text = "Get .html only";
            this.html.UseVisualStyleBackColor = true;
            this.html.Click += new System.EventHandler(this.mergeButton_Click);
            // 
            // excelToo
            // 
            this.excelToo.Location = new System.Drawing.Point(414, 680);
            this.excelToo.Name = "excelToo";
            this.excelToo.Size = new System.Drawing.Size(160, 50);
            this.excelToo.TabIndex = 16;
            this.excelToo.Text = "Get both .html and .xlsx";
            this.excelToo.UseVisualStyleBackColor = true;
            this.excelToo.Click += new System.EventHandler(this.mergeButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ScrollBar;
            this.ClientSize = new System.Drawing.Size(730, 777);
            this.Controls.Add(this.excelToo);
            this.Controls.Add(this.html);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.description);
            this.Controls.Add(this.title);
            this.Name = "Form1";
            this.Text = "HTML FileMerger";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label title;
        private System.Windows.Forms.Label description;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button html;
        private System.Windows.Forms.Button excelToo;
    }
}

