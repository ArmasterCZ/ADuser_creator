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
            this.excelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.TS_MenuItem_loadExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.ts_TextBox1 = new System.Windows.Forms.ToolStripTextBox();
            this.testToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.test1ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.TS_getPath = new System.Windows.Forms.ToolStripMenuItem();
            this.ts_TextBox2 = new System.Windows.Forms.ToolStripTextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTable1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // bWrite
            // 
            this.bWrite.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.bWrite.Location = new System.Drawing.Point(1324, 171);
            this.bWrite.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.bWrite.Name = "bWrite";
            this.bWrite.Size = new System.Drawing.Size(100, 28);
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
            this.dataGridView1.Location = new System.Drawing.Point(16, 33);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1408, 130);
            this.dataGridView1.TabIndex = 2;
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
            this.label_Actual.Location = new System.Drawing.Point(16, 177);
            this.label_Actual.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label_Actual.Name = "label_Actual";
            this.label_Actual.Size = new System.Drawing.Size(20, 17);
            this.label_Actual.TabIndex = 4;
            this.label_Actual.Text = "...";
            // 
            // bSearch
            // 
            this.bSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.bSearch.Location = new System.Drawing.Point(1216, 171);
            this.bSearch.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.bSearch.Name = "bSearch";
            this.bSearch.Size = new System.Drawing.Size(100, 28);
            this.bSearch.TabIndex = 0;
            this.bSearch.Text = "Search";
            this.bSearch.UseVisualStyleBackColor = true;
            this.bSearch.Click += new System.EventHandler(this.bSearch_Click);
            // 
            // bDelete
            // 
            this.bDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bDelete.Location = new System.Drawing.Point(737, 171);
            this.bDelete.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.bDelete.Name = "bDelete";
            this.bDelete.Size = new System.Drawing.Size(100, 28);
            this.bDelete.TabIndex = 0;
            this.bDelete.Text = "Clean Table";
            this.bDelete.UseVisualStyleBackColor = true;
            this.bDelete.Click += new System.EventHandler(this.bDelete_Click);
            // 
            // bClone
            // 
            this.bClone.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bClone.Location = new System.Drawing.Point(629, 171);
            this.bClone.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.bClone.Name = "bClone";
            this.bClone.Size = new System.Drawing.Size(100, 28);
            this.bClone.TabIndex = 0;
            this.bClone.Text = "Clone";
            this.bClone.UseVisualStyleBackColor = true;
            this.bClone.Click += new System.EventHandler(this.bClone_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.excelToolStripMenuItem,
            this.testToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(8, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(1440, 28);
            this.menuStrip1.TabIndex = 5;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // excelToolStripMenuItem
            // 
            this.excelToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.TS_MenuItem_loadExcel,
            this.ts_TextBox1,
            this.TS_getPath,
            this.ts_TextBox2});
            this.excelToolStripMenuItem.Name = "excelToolStripMenuItem";
            this.excelToolStripMenuItem.Size = new System.Drawing.Size(55, 24);
            this.excelToolStripMenuItem.Text = "Excel";
            // 
            // TS_MenuItem_loadExcel
            // 
            this.TS_MenuItem_loadExcel.Name = "TS_MenuItem_loadExcel";
            this.TS_MenuItem_loadExcel.Size = new System.Drawing.Size(186, 26);
            this.TS_MenuItem_loadExcel.Text = "Load Excel line:";
            this.TS_MenuItem_loadExcel.Click += new System.EventHandler(this.TS_MenuItem_loadExcel_Click);
            // 
            // ts_TextBox1
            // 
            this.ts_TextBox1.Name = "ts_TextBox1";
            this.ts_TextBox1.Size = new System.Drawing.Size(100, 27);
            this.ts_TextBox1.Text = "1";
            // 
            // testToolStripMenuItem
            // 
            this.testToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.test1ToolStripMenuItem});
            this.testToolStripMenuItem.Name = "testToolStripMenuItem";
            this.testToolStripMenuItem.Size = new System.Drawing.Size(47, 24);
            this.testToolStripMenuItem.Text = "Test";
            // 
            // test1ToolStripMenuItem
            // 
            this.test1ToolStripMenuItem.Name = "test1ToolStripMenuItem";
            this.test1ToolStripMenuItem.Size = new System.Drawing.Size(181, 26);
            this.test1ToolStripMenuItem.Text = "test1";
            this.test1ToolStripMenuItem.Click += new System.EventHandler(this.bTest_Click);
            // 
            // TS_getPath
            // 
            this.TS_getPath.Name = "TS_getPath";
            this.TS_getPath.Size = new System.Drawing.Size(191, 26);
            this.TS_getPath.Text = "Get current path";
            this.TS_getPath.Click += new System.EventHandler(this.TS_getPath_Click);
            // 
            // ts_TextBox2
            // 
            this.ts_TextBox2.AcceptsReturn = true;
            this.ts_TextBox2.Name = "ts_TextBox2";
            this.ts_TextBox2.Size = new System.Drawing.Size(100, 27);
            this.ts_TextBox2.Text = "C:/info.xls";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1440, 214);
            this.Controls.Add(this.label_Actual);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.bClone);
            this.Controls.Add(this.bDelete);
            this.Controls.Add(this.bSearch);
            this.Controls.Add(this.bWrite);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "Form1";
            this.Text = "PowerShell - AD User Creator";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTable1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
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
        private System.Windows.Forms.ToolStripMenuItem excelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem TS_MenuItem_loadExcel;
        private System.Windows.Forms.ToolStripTextBox ts_TextBox1;
        private System.Windows.Forms.ToolStripMenuItem testToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem test1ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem TS_getPath;
        private System.Windows.Forms.ToolStripTextBox ts_TextBox2;
    }
}

