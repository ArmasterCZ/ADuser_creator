namespace ADuser_creator
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.b_Write = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dataSet1 = new System.Data.DataSet();
            this.dataColumn1 = new System.Data.DataColumn();
            this.dataColumn2 = new System.Data.DataColumn();
            this.dataColumn3 = new System.Data.DataColumn();
            this.dataColumn4 = new System.Data.DataColumn();
            this.label_Actual = new System.Windows.Forms.Label();
            this.b_Search = new System.Windows.Forms.Button();
            this.b_Delete = new System.Windows.Forms.Button();
            this.b_Clone = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.ts_excel = new System.Windows.Forms.ToolStripMenuItem();
            this.ts_MenuItem_loadExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.ts_TextBox1 = new System.Windows.Forms.ToolStripTextBox();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.ts_changePathExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.ts_ExTextBoxPath = new System.Windows.Forms.ToolStripTextBox();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.ts_getPathAC = new System.Windows.Forms.ToolStripMenuItem();
            this.ts_ExTextBoxPathAC = new System.Windows.Forms.ToolStripTextBox();
            this.ts_userSetting = new System.Windows.Forms.ToolStripMenuItem();
            this.ts_MenuItem_moveUser = new System.Windows.Forms.ToolStripMenuItem();
            this.ts_test = new System.Windows.Forms.ToolStripMenuItem();
            this.ts_test1 = new System.Windows.Forms.ToolStripMenuItem();
            this.nic1ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripTextBox1 = new System.Windows.Forms.ToolStripTextBox();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.nic2ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripComboBox1 = new System.Windows.Forms.ToolStripComboBox();
            this.ts_createTestUser = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // b_Write
            // 
            this.b_Write.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.b_Write.Location = new System.Drawing.Point(430, 542);
            this.b_Write.Name = "b_Write";
            this.b_Write.Size = new System.Drawing.Size(75, 23);
            this.b_Write.TabIndex = 0;
            this.b_Write.Text = "Write";
            this.b_Write.UseVisualStyleBackColor = true;
            this.b_Write.Click += new System.EventHandler(this.b_Write_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowDrop = true;
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 27);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(493, 509);
            this.dataGridView1.TabIndex = 2;
            this.dataGridView1.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellEnter);
            this.dataGridView1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.dataGridView1_KeyUp);
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "NewDataSet";
            // 
            // dataColumn1
            // 
            this.dataColumn1.ColumnName = "ID";
            // 
            // dataColumn2
            // 
            this.dataColumn2.ColumnName = "Name";
            // 
            // dataColumn3
            // 
            this.dataColumn3.ColumnName = "Info1";
            // 
            // dataColumn4
            // 
            this.dataColumn4.ColumnName = "Info2";
            // 
            // label_Actual
            // 
            this.label_Actual.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label_Actual.AutoSize = true;
            this.label_Actual.Location = new System.Drawing.Point(12, 547);
            this.label_Actual.Name = "label_Actual";
            this.label_Actual.Size = new System.Drawing.Size(16, 13);
            this.label_Actual.TabIndex = 4;
            this.label_Actual.Text = "...";
            // 
            // b_Search
            // 
            this.b_Search.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.b_Search.Location = new System.Drawing.Point(349, 542);
            this.b_Search.Name = "b_Search";
            this.b_Search.Size = new System.Drawing.Size(75, 23);
            this.b_Search.TabIndex = 0;
            this.b_Search.Text = "Search";
            this.b_Search.UseVisualStyleBackColor = true;
            this.b_Search.Click += new System.EventHandler(this.b_Search_Click);
            // 
            // b_Delete
            // 
            this.b_Delete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.b_Delete.Location = new System.Drawing.Point(268, 542);
            this.b_Delete.Name = "b_Delete";
            this.b_Delete.Size = new System.Drawing.Size(75, 23);
            this.b_Delete.TabIndex = 0;
            this.b_Delete.Text = "Clean Table";
            this.b_Delete.UseVisualStyleBackColor = true;
            this.b_Delete.Click += new System.EventHandler(this.b_Delete_Click);
            // 
            // b_Clone
            // 
            this.b_Clone.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.b_Clone.Location = new System.Drawing.Point(187, 542);
            this.b_Clone.Name = "b_Clone";
            this.b_Clone.Size = new System.Drawing.Size(75, 23);
            this.b_Clone.TabIndex = 0;
            this.b_Clone.Text = "Clone";
            this.b_Clone.UseVisualStyleBackColor = true;
            this.b_Clone.Click += new System.EventHandler(this.b_Clone_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ts_excel,
            this.ts_userSetting,
            this.ts_test});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(517, 24);
            this.menuStrip1.TabIndex = 5;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // ts_excel
            // 
            this.ts_excel.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ts_MenuItem_loadExcel,
            this.ts_TextBox1,
            this.toolStripSeparator2,
            this.ts_changePathExcel,
            this.ts_ExTextBoxPath,
            this.toolStripSeparator3,
            this.ts_getPathAC,
            this.ts_ExTextBoxPathAC});
            this.ts_excel.Name = "ts_excel";
            this.ts_excel.Size = new System.Drawing.Size(45, 20);
            this.ts_excel.Text = "Excel";
            // 
            // ts_MenuItem_loadExcel
            // 
            this.ts_MenuItem_loadExcel.Name = "ts_MenuItem_loadExcel";
            this.ts_MenuItem_loadExcel.Size = new System.Drawing.Size(223, 22);
            this.ts_MenuItem_loadExcel.Text = "Load Excel line:";
            this.ts_MenuItem_loadExcel.Click += new System.EventHandler(this.ts_loadExcel_Click);
            // 
            // ts_TextBox1
            // 
            this.ts_TextBox1.Name = "ts_TextBox1";
            this.ts_TextBox1.Size = new System.Drawing.Size(100, 23);
            this.ts_TextBox1.Text = "1";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(220, 6);
            // 
            // ts_changePathExcel
            // 
            this.ts_changePathExcel.Name = "ts_changePathExcel";
            this.ts_changePathExcel.Size = new System.Drawing.Size(223, 22);
            this.ts_changePathExcel.Text = "Change Excel Path";
            this.ts_changePathExcel.Click += new System.EventHandler(this.ts_changePathExcel_Click);
            // 
            // ts_ExTextBoxPath
            // 
            this.ts_ExTextBoxPath.AcceptsReturn = true;
            this.ts_ExTextBoxPath.Enabled = false;
            this.ts_ExTextBoxPath.Name = "ts_ExTextBoxPath";
            this.ts_ExTextBoxPath.Size = new System.Drawing.Size(100, 23);
            this.ts_ExTextBoxPath.Text = "D:/info.xlsx";
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(220, 6);
            // 
            // ts_getPathAC
            // 
            this.ts_getPathAC.Name = "ts_getPathAC";
            this.ts_getPathAC.Size = new System.Drawing.Size(223, 22);
            this.ts_getPathAC.Text = "Current path for completion";
            this.ts_getPathAC.Click += new System.EventHandler(this.ts_getPathAC_Click);
            // 
            // ts_ExTextBoxPathAC
            // 
            this.ts_ExTextBoxPathAC.Name = "ts_ExTextBoxPathAC";
            this.ts_ExTextBoxPathAC.Size = new System.Drawing.Size(100, 23);
            this.ts_ExTextBoxPathAC.Text = "D:/Office.xlsx";
            // 
            // ts_userSetting
            // 
            this.ts_userSetting.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ts_MenuItem_moveUser});
            this.ts_userSetting.Name = "ts_userSetting";
            this.ts_userSetting.Size = new System.Drawing.Size(82, 20);
            this.ts_userSetting.Text = "User Setting";
            // 
            // ts_MenuItem_moveUser
            // 
            this.ts_MenuItem_moveUser.Name = "ts_MenuItem_moveUser";
            this.ts_MenuItem_moveUser.Size = new System.Drawing.Size(152, 22);
            this.ts_MenuItem_moveUser.Text = "Move User";
            this.ts_MenuItem_moveUser.Click += new System.EventHandler(this.ts_moveUser_Click);
            // 
            // ts_test
            // 
            this.ts_test.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ts_test1,
            this.ts_createTestUser});
            this.ts_test.Name = "ts_test";
            this.ts_test.Size = new System.Drawing.Size(41, 20);
            this.ts_test.Text = "Test";
            // 
            // ts_test1
            // 
            this.ts_test1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.nic1ToolStripMenuItem,
            this.toolStripTextBox1,
            this.toolStripSeparator1,
            this.nic2ToolStripMenuItem,
            this.toolStripComboBox1});
            this.ts_test1.Name = "ts_test1";
            this.ts_test1.Size = new System.Drawing.Size(158, 22);
            this.ts_test1.Text = "test (nic nedělá)";
            // 
            // nic1ToolStripMenuItem
            // 
            this.nic1ToolStripMenuItem.Name = "nic1ToolStripMenuItem";
            this.nic1ToolStripMenuItem.Size = new System.Drawing.Size(181, 22);
            this.nic1ToolStripMenuItem.Text = "nic1";
            // 
            // toolStripTextBox1
            // 
            this.toolStripTextBox1.Name = "toolStripTextBox1";
            this.toolStripTextBox1.Size = new System.Drawing.Size(100, 23);
            this.toolStripTextBox1.Text = "Bla Bla";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(178, 6);
            // 
            // nic2ToolStripMenuItem
            // 
            this.nic2ToolStripMenuItem.Name = "nic2ToolStripMenuItem";
            this.nic2ToolStripMenuItem.Size = new System.Drawing.Size(181, 22);
            this.nic2ToolStripMenuItem.Text = "nic2";
            // 
            // toolStripComboBox1
            // 
            this.toolStripComboBox1.Items.AddRange(new object[] {
            "Test1",
            "Test2",
            "Test3",
            "TTtt",
            "TT-",
            "AA+",
            "aa-",
            "aaa1"});
            this.toolStripComboBox1.Name = "toolStripComboBox1";
            this.toolStripComboBox1.Size = new System.Drawing.Size(121, 23);
            // 
            // ts_createTestUser
            // 
            this.ts_createTestUser.Name = "ts_createTestUser";
            this.ts_createTestUser.Size = new System.Drawing.Size(158, 22);
            this.ts_createTestUser.Text = "create test user";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(517, 577);
            this.Controls.Add(this.label_Actual);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.b_Clone);
            this.Controls.Add(this.b_Delete);
            this.Controls.Add(this.b_Search);
            this.Controls.Add(this.b_Write);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "PowerShell - AD User Creator";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button b_Write;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Data.DataSet dataSet1;
        private System.Data.DataColumn dataColumn1;
        private System.Data.DataColumn dataColumn2;
        private System.Data.DataColumn dataColumn3;
        private System.Data.DataColumn dataColumn4;
        private System.Windows.Forms.Label label_Actual;
        private System.Windows.Forms.Button b_Search;
        private System.Windows.Forms.Button b_Delete;
        private System.Windows.Forms.Button b_Clone;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem ts_excel;
        private System.Windows.Forms.ToolStripMenuItem ts_MenuItem_loadExcel;
        private System.Windows.Forms.ToolStripTextBox ts_TextBox1;
        private System.Windows.Forms.ToolStripMenuItem ts_test;
        private System.Windows.Forms.ToolStripMenuItem ts_test1;
        private System.Windows.Forms.ToolStripMenuItem ts_changePathExcel;
        private System.Windows.Forms.ToolStripTextBox ts_ExTextBoxPath;
        private System.Windows.Forms.ToolStripMenuItem ts_userSetting;
        private System.Windows.Forms.ToolStripMenuItem ts_MenuItem_moveUser;
        private System.Windows.Forms.ToolStripMenuItem ts_createTestUser;
        private System.Windows.Forms.ToolStripMenuItem nic1ToolStripMenuItem;
        private System.Windows.Forms.ToolStripTextBox toolStripTextBox1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem nic2ToolStripMenuItem;
        private System.Windows.Forms.ToolStripComboBox toolStripComboBox1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripTextBox ts_ExTextBoxPathAC;
        private System.Windows.Forms.ToolStripMenuItem ts_getPathAC;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
    }
}

