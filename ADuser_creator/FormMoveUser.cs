using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADuser_creator
{
    public partial class FormMoveUser : Form
    {
        public FormMoveUser(string userName)
        {
            InitializeComponent();
            //set start position of form
            this.StartPosition = FormStartPosition.CenterParent;
            this.userName = userName;
            bMove.Text = $"Přesunout {userName}";
        }

        private string userName;
        public string resultText { get; set; }

        private void bClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void bMove_Click(object sender, EventArgs e)
        {
            resultText = cbPath.Text;
            DialogResult boxResult = MessageBox.Show($"Opravdu chcete přesunout uživatele {userName} do kontejneru Environment{Environment.NewLine}{resultText}",
                "Přesunout uživatele?",MessageBoxButtons.YesNo,MessageBoxIcon.Asterisk);
            if (boxResult == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.Yes;
                Close();
            }
        }

        private void bComfirm_Click(object sender, EventArgs e)
        {
            resultText = cbPath.Text;
            this.DialogResult = DialogResult.No;
            Close();
        }
    }
}
