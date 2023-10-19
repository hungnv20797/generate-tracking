namespace GenerateTracking
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
            this.btnExportYO = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnExportYO
            // 
            this.btnExportYO.Location = new System.Drawing.Point(12, 26);
            this.btnExportYO.Name = "btnExportYO";
            this.btnExportYO.Size = new System.Drawing.Size(321, 50);
            this.btnExportYO.TabIndex = 0;
            this.btnExportYO.Text = "Export YO";
            this.btnExportYO.UseVisualStyleBackColor = true;
            this.btnExportYO.Click += new System.EventHandler(this.btnExportYO_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(345, 446);
            this.Controls.Add(this.btnExportYO);
            this.Name = "Form1";
            this.Text = "Generate Tracking";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnExportYO;
    }
}

