namespace GeraExcelEppPlus
{
    partial class Main
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
            this.btnGeraExcel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnGeraExcel
            // 
            this.btnGeraExcel.Location = new System.Drawing.Point(38, 12);
            this.btnGeraExcel.Name = "btnGeraExcel";
            this.btnGeraExcel.Size = new System.Drawing.Size(148, 23);
            this.btnGeraExcel.TabIndex = 0;
            this.btnGeraExcel.Text = "Gerar Excel de Exemplo";
            this.btnGeraExcel.UseVisualStyleBackColor = true;
            this.btnGeraExcel.Click += new System.EventHandler(this.btnGeraExcel_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(235, 56);
            this.Controls.Add(this.btnGeraExcel);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Main";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gera Excel com EPPPLUS";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnGeraExcel;
    }
}

