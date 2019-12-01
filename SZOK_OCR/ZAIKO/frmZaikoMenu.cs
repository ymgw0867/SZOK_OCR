using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SZOK_OCR.ZAIKO
{
    public partial class frmZaikoMenu : Form
    {
        public frmZaikoMenu()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmShukkoImport frm = new frmShukkoImport();
            frm.ShowDialog();
            this.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmKaishuData frm = new frmKaishuData();
            frm.ShowDialog();
            this.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmZaikoSum frm = new frmZaikoSum();
            frm.ShowDialog();
            this.Show();
        }
    }
}
