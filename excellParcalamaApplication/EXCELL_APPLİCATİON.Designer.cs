namespace excellParcalamaApplication
{
    partial class EXCELL_APPLİCATİON
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
            this.components = new System.ComponentModel.Container();
            this.button_parcala = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox_colomn_name = new System.Windows.Forms.TextBox();
            this.button_exit = new System.Windows.Forms.Button();
            this.mailBody = new System.Windows.Forms.TextBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.SuspendLayout();
            // 
            // button_parcala
            // 
            this.button_parcala.BackColor = System.Drawing.Color.LightGray;
            this.button_parcala.ForeColor = System.Drawing.Color.DarkRed;
            this.button_parcala.Location = new System.Drawing.Point(12, 65);
            this.button_parcala.Name = "button_parcala";
            this.button_parcala.Size = new System.Drawing.Size(275, 58);
            this.button_parcala.TabIndex = 0;
            this.button_parcala.Text = "EMAIL GONDER";
            this.button_parcala.UseVisualStyleBackColor = false;
            this.button_parcala.Click += new System.EventHandler(this.button_parcala_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(124, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "SÜTUN İSMİNİ GİRİNİZ";
            // 
            // textBox_colomn_name
            // 
            this.textBox_colomn_name.Location = new System.Drawing.Point(142, 28);
            this.textBox_colomn_name.Name = "textBox_colomn_name";
            this.textBox_colomn_name.Size = new System.Drawing.Size(145, 20);
            this.textBox_colomn_name.TabIndex = 2;
            // 
            // button_exit
            // 
            this.button_exit.BackColor = System.Drawing.Color.LightGray;
            this.button_exit.ForeColor = System.Drawing.Color.DarkRed;
            this.button_exit.Location = new System.Drawing.Point(12, 129);
            this.button_exit.Name = "button_exit";
            this.button_exit.Size = new System.Drawing.Size(275, 58);
            this.button_exit.TabIndex = 3;
            this.button_exit.Text = "ÇIKIŞ";
            this.button_exit.UseVisualStyleBackColor = false;
            this.button_exit.Click += new System.EventHandler(this.button_exit_Click);
            // 
            // mailBody
            // 
            this.mailBody.Location = new System.Drawing.Point(12, 243);
            this.mailBody.Multiline = true;
            this.mailBody.Name = "mailBody";
            this.mailBody.Size = new System.Drawing.Size(275, 233);
            this.mailBody.TabIndex = 4;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // EXCELL_APPLİCATİON
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(299, 486);
            this.Controls.Add(this.mailBody);
            this.Controls.Add(this.button_exit);
            this.Controls.Add(this.textBox_colomn_name);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button_parcala);
            this.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.Name = "EXCELL_APPLİCATİON";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.EXCELL_APPLİCATİON_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button_parcala;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_colomn_name;
        private System.Windows.Forms.Button button_exit;
        private System.Windows.Forms.TextBox mailBody;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
    }
}

