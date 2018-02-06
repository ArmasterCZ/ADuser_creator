namespace ADuser_creator
{
    partial class FormMoveUser
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
            this.cbPath = new System.Windows.Forms.ComboBox();
            this.bMove = new System.Windows.Forms.Button();
            this.bComfirm = new System.Windows.Forms.Button();
            this.bClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cbPath
            // 
            this.cbPath.FormattingEnabled = true;
            this.cbPath.Items.AddRange(new object[] {
            "OU=Users,OU=People,OU=Company,DC=sitel,DC=cz",
            "OU=Test,OU=Service,OU=Company,DC=sitel,DC=cz",
            "OU=Only Contacts,OU=People,OU=Company,DC=sitel,DC=cz"});
            this.cbPath.Location = new System.Drawing.Point(13, 13);
            this.cbPath.Name = "cbPath";
            this.cbPath.Size = new System.Drawing.Size(446, 21);
            this.cbPath.TabIndex = 0;
            // 
            // bMove
            // 
            this.bMove.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.bMove.Location = new System.Drawing.Point(180, 40);
            this.bMove.Name = "bMove";
            this.bMove.Size = new System.Drawing.Size(157, 23);
            this.bMove.TabIndex = 1;
            this.bMove.Text = "Přesunout";
            this.bMove.UseVisualStyleBackColor = true;
            this.bMove.Click += new System.EventHandler(this.bMove_Click);
            // 
            // bComfirm
            // 
            this.bComfirm.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.bComfirm.Location = new System.Drawing.Point(343, 40);
            this.bComfirm.Name = "bComfirm";
            this.bComfirm.Size = new System.Drawing.Size(116, 23);
            this.bComfirm.TabIndex = 1;
            this.bComfirm.Text = "Zapsat do tabulky";
            this.bComfirm.UseVisualStyleBackColor = true;
            this.bComfirm.Click += new System.EventHandler(this.bComfirm_Click);
            // 
            // bClose
            // 
            this.bClose.Location = new System.Drawing.Point(13, 40);
            this.bClose.Name = "bClose";
            this.bClose.Size = new System.Drawing.Size(75, 23);
            this.bClose.TabIndex = 1;
            this.bClose.Text = "Zavři";
            this.bClose.UseVisualStyleBackColor = true;
            this.bClose.Click += new System.EventHandler(this.bClose_Click);
            // 
            // FormMoveUser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(471, 74);
            this.Controls.Add(this.bComfirm);
            this.Controls.Add(this.bClose);
            this.Controls.Add(this.bMove);
            this.Controls.Add(this.cbPath);
            this.Name = "FormMoveUser";
            this.Text = "Správa kontejneru";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox cbPath;
        private System.Windows.Forms.Button bMove;
        private System.Windows.Forms.Button bComfirm;
        private System.Windows.Forms.Button bClose;
    }
}