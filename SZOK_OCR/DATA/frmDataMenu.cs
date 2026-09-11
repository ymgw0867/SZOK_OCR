using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SZOK_OCR.Common;

namespace SZOK_OCR.DATA
{
    public partial class frmDataMenu : Form
    {
        public frmDataMenu(DataGridViewRow r)
        {
            InitializeComponent();
            _row = r;
        }

        int _ix;

        DataGridViewRow _row;

        private void frmDataMenu_Load(object sender, EventArgs e)
        {
            lblNumber.Text = _row.Cells[1].Value.ToString();
            lblName.Text = _row.Cells[10].Value.ToString();
            lblAddress.Text = _row.Cells[9].Value.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 指定データとカード画像を表示
            ShowPastData(_ix);
        }

        private void ShowPastData(int iX)
        {
            this.Hide();
            using (frmPastData frm = new frmPastData(iX))
            {
                frm.ShowDialog();
                this.Show();
                //EditStatus = frm.EditMode;
            }
        }
    }
}
