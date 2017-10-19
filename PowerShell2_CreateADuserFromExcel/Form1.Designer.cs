namespace PowerShell2_CreateADuserFromExcel
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.bWrite = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dataSet1 = new System.Data.DataSet();
            this.dataTable1 = new System.Data.DataTable();
            this.dataColumn1 = new System.Data.DataColumn();
            this.dataColumn2 = new System.Data.DataColumn();
            this.dataColumn3 = new System.Data.DataColumn();
            this.dataColumn4 = new System.Data.DataColumn();
            this.label_Actual = new System.Windows.Forms.Label();
            this.bSearch = new System.Windows.Forms.Button();
            this.bDelete = new System.Windows.Forms.Button();
            this.bClone = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.TS_excel = new System.Windows.Forms.ToolStripMenuItem();
            this.TS_MenuItem_loadExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.ts_TextBox1 = new System.Windows.Forms.ToolStripTextBox();
            this.TS_getPath = new System.Windows.Forms.ToolStripMenuItem();
            this.ts_TextBox2 = new System.Windows.Forms.ToolStripTextBox();
            this.TS_test = new System.Windows.Forms.ToolStripMenuItem();
            this.TS_test1 = new System.Windows.Forms.ToolStripMenuItem();
            this.TS_createTestUser = new System.Windows.Forms.ToolStripMenuItem();
            this.TS_userSetting = new System.Windows.Forms.ToolStripMenuItem();
            this.TS_MenuItem_moveUser = new System.Windows.Forms.ToolStripMenuItem();
            this.ts_TextBox3 = new System.Windows.Forms.ToolStripTextBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.userContainerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.testContainerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.clearToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTable1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // bWrite
            // 
            this.bWrite.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.bWrite.Location = new System.Drawing.Point(993, 139);
            this.bWrite.Name = "bWrite";
            this.bWrite.Size = new System.Drawing.Size(75, 23);
            this.bWrite.TabIndex = 0;
            this.bWrite.Text = "Write";
            this.bWrite.UseVisualStyleBackColor = true;
            this.bWrite.Click += new System.EventHandler(this.bWrite_Click);
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
            this.dataGridView1.Size = new System.Drawing.Size(1056, 106);
            this.dataGridView1.TabIndex = 2;
            this.dataGridView1.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellEnter);
            this.dataGridView1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.dataGridView1_KeyUp);
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "NewDataSet";
            this.dataSet1.Tables.AddRange(new System.Data.DataTable[] {
            this.dataTable1});
            // 
            // dataTable1
            // 
            this.dataTable1.Columns.AddRange(new System.Data.DataColumn[] {
            this.dataColumn1,
            this.dataColumn2,
            this.dataColumn3,
            this.dataColumn4});
            this.dataTable1.TableName = "Table1";
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
            this.label_Actual.Location = new System.Drawing.Point(12, 144);
            this.label_Actual.Name = "label_Actual";
            this.label_Actual.Size = new System.Drawing.Size(16, 13);
            this.label_Actual.TabIndex = 4;
            this.label_Actual.Text = "...";
            // 
            // bSearch
            // 
            this.bSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.bSearch.Location = new System.Drawing.Point(912, 139);
            this.bSearch.Name = "bSearch";
            this.bSearch.Size = new System.Drawing.Size(75, 23);
            this.bSearch.TabIndex = 0;
            this.bSearch.Text = "Search";
            this.bSearch.UseVisualStyleBackColor = true;
            this.bSearch.Click += new System.EventHandler(this.bSearch_Click);
            // 
            // bDelete
            // 
            this.bDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bDelete.Location = new System.Drawing.Point(553, 139);
            this.bDelete.Name = "bDelete";
            this.bDelete.Size = new System.Drawing.Size(75, 23);
            this.bDelete.TabIndex = 0;
            this.bDelete.Text = "Clean Table";
            this.bDelete.UseVisualStyleBackColor = true;
            this.bDelete.Click += new System.EventHandler(this.bDelete_Click);
            // 
            // bClone
            // 
            this.bClone.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bClone.Location = new System.Drawing.Point(472, 139);
            this.bClone.Name = "bClone";
            this.bClone.Size = new System.Drawing.Size(75, 23);
            this.bClone.TabIndex = 0;
            this.bClone.Text = "Clone";
            this.bClone.UseVisualStyleBackColor = true;
            this.bClone.Click += new System.EventHandler(this.bClone_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.TS_excel,
            this.TS_test,
            this.TS_userSetting});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1080, 24);
            this.menuStrip1.TabIndex = 5;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // TS_excel
            // 
            this.TS_excel.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.TS_MenuItem_loadExcel,
            this.ts_TextBox1,
            this.TS_getPath,
            this.ts_TextBox2});
            this.TS_excel.Name = "TS_excel";
            this.TS_excel.Size = new System.Drawing.Size(45, 20);
            this.TS_excel.Text = "Excel";
            // 
            // TS_MenuItem_loadExcel
            // 
            this.TS_MenuItem_loadExcel.Name = "TS_MenuItem_loadExcel";
            this.TS_MenuItem_loadExcel.Size = new System.Drawing.Size(160, 22);
            this.TS_MenuItem_loadExcel.Text = "Load Excel line:";
            this.TS_MenuItem_loadExcel.Click += new System.EventHandler(this.TS_MenuItem_loadExcel_Click);
            // 
            // ts_TextBox1
            // 
            this.ts_TextBox1.Name = "ts_TextBox1";
            this.ts_TextBox1.Size = new System.Drawing.Size(100, 23);
            this.ts_TextBox1.Text = "1";
            // 
            // TS_getPath
            // 
            this.TS_getPath.Name = "TS_getPath";
            this.TS_getPath.Size = new System.Drawing.Size(160, 22);
            this.TS_getPath.Text = "Get current path";
            this.TS_getPath.Click += new System.EventHandler(this.TS_getPath_Click);
            // 
            // ts_TextBox2
            // 
            this.ts_TextBox2.AcceptsReturn = true;
            this.ts_TextBox2.Name = "ts_TextBox2";
            this.ts_TextBox2.Size = new System.Drawing.Size(100, 23);
            this.ts_TextBox2.Text = "D:/info.xls";
            // 
            // TS_test
            // 
            this.TS_test.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.TS_test1,
            this.TS_createTestUser});
            this.TS_test.Name = "TS_test";
            this.TS_test.Size = new System.Drawing.Size(41, 20);
            this.TS_test.Text = "Test";
            // 
            // TS_test1
            // 
            this.TS_test1.Name = "TS_test1";
            this.TS_test1.Size = new System.Drawing.Size(153, 22);
            this.TS_test1.Text = "test1";
            this.TS_test1.Click += new System.EventHandler(this.bTest_Click);
            // 
            // TS_createTestUser
            // 
            this.TS_createTestUser.Name = "TS_createTestUser";
            this.TS_createTestUser.Size = new System.Drawing.Size(153, 22);
            this.TS_createTestUser.Text = "create test user";
            this.TS_createTestUser.Click += new System.EventHandler(this.TS_createTestUser_Click);
            // 
            // TS_userSetting
            // 
            this.TS_userSetting.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.TS_MenuItem_moveUser,
            this.ts_TextBox3});
            this.TS_userSetting.Name = "TS_userSetting";
            this.TS_userSetting.Size = new System.Drawing.Size(82, 20);
            this.TS_userSetting.Text = "User Setting";
            // 
            // TS_MenuItem_moveUser
            // 
            this.TS_MenuItem_moveUser.Name = "TS_MenuItem_moveUser";
            this.TS_MenuItem_moveUser.Size = new System.Drawing.Size(360, 22);
            this.TS_MenuItem_moveUser.Text = "Move User";
            this.TS_MenuItem_moveUser.Click += new System.EventHandler(this.TS_MenuItem_moveUser_Click);
            // 
            // ts_TextBox3
            // 
            this.ts_TextBox3.Name = "ts_TextBox3";
            this.ts_TextBox3.Size = new System.Drawing.Size(300, 23);
            this.ts_TextBox3.Text = "OU=Test,OU=Service,OU=Company,DC=sitel,DC=cz";
            this.ts_TextBox3.MouseDown += new System.Windows.Forms.MouseEventHandler(this.ts_TextBox3_MouseDown);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.userContainerToolStripMenuItem,
            this.testContainerToolStripMenuItem,
            this.clearToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(153, 70);
            // 
            // userContainerToolStripMenuItem
            // 
            this.userContainerToolStripMenuItem.Name = "userContainerToolStripMenuItem";
            this.userContainerToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.userContainerToolStripMenuItem.Text = "User Container";
            this.userContainerToolStripMenuItem.Click += new System.EventHandler(this.userContainerToolStripMenuItem_Click);
            // 
            // testContainerToolStripMenuItem
            // 
            this.testContainerToolStripMenuItem.Name = "testContainerToolStripMenuItem";
            this.testContainerToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.testContainerToolStripMenuItem.Text = "Test Container";
            this.testContainerToolStripMenuItem.Click += new System.EventHandler(this.testContainerToolStripMenuItem_Click);
            // 
            // clearToolStripMenuItem
            // 
            this.clearToolStripMenuItem.Name = "clearToolStripMenuItem";
            this.clearToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.clearToolStripMenuItem.Text = "Clear";
            this.clearToolStripMenuItem.Click += new System.EventHandler(this.clearToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1080, 174);
            this.Controls.Add(this.label_Actual);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.bClone);
            this.Controls.Add(this.bDelete);
            this.Controls.Add(this.bSearch);
            this.Controls.Add(this.bWrite);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "PowerShell - AD User Creator";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTable1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bWrite;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Data.DataSet dataSet1;
        private System.Data.DataTable dataTable1;
        private System.Data.DataColumn dataColumn1;
        private System.Data.DataColumn dataColumn2;
        private System.Data.DataColumn dataColumn3;
        private System.Data.DataColumn dataColumn4;
        private System.Windows.Forms.Label label_Actual;
        private System.Windows.Forms.Button bSearch;
        private System.Windows.Forms.Button bDelete;
        private System.Windows.Forms.Button bClone;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem TS_excel;
        private System.Windows.Forms.ToolStripMenuItem TS_MenuItem_loadExcel;
        private System.Windows.Forms.ToolStripTextBox ts_TextBox1;
        private System.Windows.Forms.ToolStripMenuItem TS_test;
        private System.Windows.Forms.ToolStripMenuItem TS_test1;
        private System.Windows.Forms.ToolStripMenuItem TS_getPath;
        private System.Windows.Forms.ToolStripTextBox ts_TextBox2;
        private System.Windows.Forms.ToolStripMenuItem TS_userSetting;
        private System.Windows.Forms.ToolStripMenuItem TS_MenuItem_moveUser;
        private System.Windows.Forms.ToolStripTextBox ts_TextBox3;
        private System.Windows.Forms.ToolStripMenuItem TS_createTestUser;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem userContainerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem testContainerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem clearToolStripMenuItem;
    }
}

