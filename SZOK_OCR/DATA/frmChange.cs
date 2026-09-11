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
    public partial class frmChange : Form
    {
        public frmChange(int ix)
        {
            InitializeComponent();
            _ix = ix;
        }

        int _ix;

        private void frmChange_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 指定されたIDのデータを表示する
        /// </summary>
        /// <param name="iX">表示するデータのID</param>
        private void ShowData(int iX)
        {
            // SQL Server接続
            var master = new ClsMaster(Properties.Settings.Default.sServerName, Properties.Settings.Default.sLogin,
                                   Properties.Settings.Default.sPass, Properties.Settings.Default.sDatabase);

            var r = master.GetData<TblRegistrationCard>(iX.ToString());

            // 情報表示
            lblNumber.Text = r.Number;
            lblDate.Text = (Utility.StrtoInt(r.AddYear) + 2000) + "年" + r.AddMonth + "月" + r.AddDay + "日";
            lblName.Text = r.Name;
            lblAddress.Text = r.Address1;
            lblAddKanji.Text = r.AddressKanji;
            lblTel1.Text = r.Mobile1;
            lblTel2.Text = r.Mobile2;
            lblTel3.Text = r.Mobile3;


            dateTimePicker1.Focus();
        }

    }
}
