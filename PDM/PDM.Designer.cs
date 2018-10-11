namespace PDM
{
    partial class PDM
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
            this.button1 = new System.Windows.Forms.Button();
            this.textBox_OutputPath = new System.Windows.Forms.TextBox();
            this.label_OutputPath = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.Browse = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(381, 87);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Generate";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox_OutputPath
            // 
            this.textBox_OutputPath.Location = new System.Drawing.Point(74, 44);
            this.textBox_OutputPath.Name = "textBox_OutputPath";
            this.textBox_OutputPath.Size = new System.Drawing.Size(292, 20);
            this.textBox_OutputPath.TabIndex = 1;
            // 
            // label_OutputPath
            // 
            this.label_OutputPath.AutoSize = true;
            this.label_OutputPath.Location = new System.Drawing.Point(4, 47);
            this.label_OutputPath.Name = "label_OutputPath";
            this.label_OutputPath.Size = new System.Drawing.Size(64, 13);
            this.label_OutputPath.TabIndex = 2;
            this.label_OutputPath.Text = "Output Path";
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.HelpRequest += new System.EventHandler(this.folderBrowserDialog1_HelpRequest);
            // 
            // Browse
            // 
            this.Browse.Location = new System.Drawing.Point(381, 42);
            this.Browse.Name = "Browse";
            this.Browse.Size = new System.Drawing.Size(75, 23);
            this.Browse.TabIndex = 3;
            this.Browse.Text = "Browse";
            this.Browse.UseVisualStyleBackColor = true;
            this.Browse.Click += new System.EventHandler(this.Browse_Click);
            // 
            // PDM
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(497, 122);
            this.Controls.Add(this.Browse);
            this.Controls.Add(this.label_OutputPath);
            this.Controls.Add(this.textBox_OutputPath);
            this.Controls.Add(this.button1);
            this.Name = "PDM";
            this.Text = "PDM Generator";
            this.Load += new System.EventHandler(this.PDM_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox_OutputPath;
        private System.Windows.Forms.Label label_OutputPath;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button Browse;
    }
}

