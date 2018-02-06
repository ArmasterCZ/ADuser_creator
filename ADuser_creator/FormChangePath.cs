using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADuser_creator
{
    public partial class FormChangePath : Form
    {
        public FormChangePath(string inputPath)
        {
            InitializeComponent();
            tbPath.Text = inputPath;
        }

        public string output { get; set; }

        private void bComfirm_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(tbPath.Text))
            {
                this.output = tbPath.Text;
                this.DialogResult = DialogResult.Yes;
                this.Close();
            }
            else
            {
                MessageBox.Show("Cesta nemůže být prázdná.");
            }
        }

        private void bCurrentPath_Click(object sender, EventArgs e)
        {
            tbPath.Text = string.Format("{0}\\{1}", Directory.GetCurrentDirectory(), "info.xlsx");
        }

        private void bClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.No;
            this.Close();
        }
    }
}
