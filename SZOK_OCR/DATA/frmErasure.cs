using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SZOK_OCR.Common;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace SZOK_OCR.DATA
{
    public partial class frmErasure : Form
    {
        public frmErasure(int id)
        {
            InitializeComponent();
            _id = id;
            status = false;
        }

        public bool status { get; set; }

        int _id;

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmErasure_Load(object sender, EventArgs e)
        {
            ShowData(_id);
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
            dateTimePicker1.Value = DateTime.Today;

            dateTimePicker1.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 確認メッセージ
            if (MessageBox.Show("抹消してよろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // 抹消処理
            var master = new ClsMaster(Properties.Settings.Default.sServerName, Properties.Settings.Default.sLogin,
                                   Properties.Settings.Default.sPass, Properties.Settings.Default.sDatabase);
            master.ErasureData(_id, dateTimePicker1.Value);

            MessageBox.Show("抹消が完了しました。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);

            status = true;
            Close();
        }
    }
}
    