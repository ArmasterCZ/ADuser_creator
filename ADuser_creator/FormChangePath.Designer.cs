namespace ADuser_creator
{
    partial class FormChangePath
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
            this.bComfirm = new System.Windows.Forms.Button();
            this.bClose = new System.Windows.Forms.Button();
            this.tbPath = new System.Windows.Forms.TextBox();
            this.bCurrentPath = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // bComfirm
            // 
            this.bComfirm.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.bComfirm.Location = new System.Drawing.Point(310, 100);
            this.bComfirm.Name = "bComfirm";
            this.bComfirm.Size = new System.Drawing.Size(75, 23);
            this.bComfirm.TabIndex = 0;
            this.bComfirm.Text = "OK";
            this.bComfirm.UseVisualStyleBackColor = true;
            this.bComfirm.Click += new System.EventHandler(this.bComfirm_Click);
            // 
            // bClose
            // 
            this.bClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bClose.Location = new System.Drawing.Point(12, 100);
            this.bClose.Name = "bClose";
            this.bClose.Size = new System.Drawing.Size(75, 23);
            this.bClose.TabIndex = 0;
            this.bClose.Text = "Zavři";
            this.bClose.UseVisualStyleBackColor = true;
            this.bClose.Click += new System.EventHandler(this.bClose_Click);
            // 
            // tbPath
            // 
            this.tbPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbPath.Location = new System.Drawing.Point(12, 13);
            this.tbPath.Name = "tbPath";
            this.tbPath.Size = new System.Drawing.Size(373, 20);
            this.tbPath.TabIndex = 1;
            // 
            // bCurrentPath
            // 
            this.bCurrentPath.Location = new System.Drawing.Point(12, 39);
            this.bCurrentPath.Name = "bCurrentPath";
            this.bCurrentPath.Size = new System.Drawing.Size(160, 23);
            this.bCurrentPath.TabIndex = 0;
            this.bCurrentPath.Text = "Získej současnou cestu";
            this.bCurrentPath.UseVisualStyleBackColor = true;
            this.bCurrentPath.Click += new System.EventHandler(this.bCurrentPath_Click);
            // 
            // FormChangePath
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(397, 135);
            this.Controls.Add(this.tbPath);
            this.Controls.Add(this.bClose);
            this.Controls.Add(this.bCurrentPath);
            this.Controls.Add(this.bComfirm);
            this.Name = "FormChangePath";
            this.Text = "Změnit cestu k excelu";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bComfirm;
        private System.Windows.Forms.Button bClose;
        private System.Windows.Forms.TextBox tbPath;
        private System.Windows.Forms.Button bCurrentPath;
    }
}