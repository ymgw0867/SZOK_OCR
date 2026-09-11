using System;
using System.Collections;
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
        string[] zipArray = null;   // 郵便番号配列

        // 変更届の変更があったかどうかのフラグ：2026/09/11
        public bool EditMode { get; set; }

        private void frmChange_Load(object sender, EventArgs e)
        {
            // 郵便番号CSV配列読み込み
            Utility.zipCsvLoad(ref zipArray);

            ChkName.Checked = false;
            ChkAddress.Checked = false;
            ChkTel.Checked = false;

            txtName.Text = "";
            txtName.Enabled = false;

            txtZipCode1.Text = "";
            txtZipCode1.Enabled = false;
            txtZipCode2.Text = "";
            txtZipCode2.Enabled = false;

            txtAdd.Text = "";
            txtAdd.Enabled = false;
            txtAddKanji.Text = "";
            txtAddKanji.Enabled = false;

            txtTel1.Text = "";
            txtTel1.Enabled = false;
            txtTel2.Text = "";
            txtTel2.Enabled = false;
            txtTel3.Text = "";
            txtTel3.Enabled = false;

            // 該当データ表示
            ShowData(_ix);

            EditMode = false;
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
            lblZipCode.Text = r.ZipCode1 + "-" + r.ZipCode2;
            lblAddress.Text = r.Address1;
            lblAddKanji.Text = r.AddressKanji;
            lblTel1.Text = r.Mobile1;
            lblTel2.Text = r.Mobile2;
            lblTel3.Text = r.Mobile3;
        }

        private void ChkName_CheckedChanged(object sender, EventArgs e)
        {
            if (ChkName.Checked)
            {
                txtName.Enabled = true;
                txtName.Focus();
            }
            else
            {
                txtName.Enabled = false;
            }
        }

        private void ChkAddress_CheckedChanged(object sender, EventArgs e)
        {
            if (ChkAddress.Checked)
            {
                txtZipCode1.Enabled = true; 
                txtZipCode2.Enabled = true;
                txtAdd.Enabled = true;
                txtAddKanji.Enabled = true;
                txtZipCode1.Focus();
            }
            else
            {
                txtZipCode1.Enabled = false; 
                txtZipCode2.Enabled = false;
                txtAdd.Enabled = false;
                txtAddKanji.Enabled = false;
            }
        }

        private void ChkTel_CheckedChanged(object sender, EventArgs e)
        {
            if (ChkTel.Checked)
            {
                txtTel1.Enabled = true;
                txtTel2.Enabled = true;
                txtTel3.Enabled = true;
                txtTel1.Focus();
            }
            else
            {
                txtTel1.Enabled = false;
                txtTel2.Enabled = false;
                txtTel3.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 変更届登録
            EditMode = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 郵便番号から住所変換：2026/09/11
            var (zipAdd, zipAddKN, iZ) = Utility.GetAddressFromZipCode(txtZipCode1.Text, txtZipCode2.Text, zipArray);

            if (iZ == 1)
            {
                // 単独で該当郵便番号あり
                if (zipAdd != string.Empty)
                {
                    // 住所フリガナ
                    txtAdd.Text = zipAdd;

                    // 住所漢字
                    //string knAdd = (txtAddKanji.Text.Replace(" ", "").Replace(zipAddKN.Replace(" ", ""), "")).Trim();
                    txtAddKanji.Text = zipAddKN;
                }
            }
            else if (iZ > 1)
            {
                // 複数の該当郵便番号あり
                string msg = "郵便番号 " + txtZipCode1.Text.Trim() + txtZipCode2.Text.Trim() + " は複数の地名が存在します。" + Environment.NewLine;
                msg += "「〒⇔住所」ボタンから該当する地名を選択してください";
                MessageBox.Show(msg, "複数地名あり", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (iZ == 0)
            {
                // 該当郵便番号なし
                string msg = "郵便番号 " + txtZipCode1.Text.Trim() + txtZipCode2.Text.Trim() + " は存在しません。" + Environment.NewLine;
                msg += "「〒⇔住所」ボタンから該当する地名を選択してください";
                MessageBox.Show(msg, "該当地名なし", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var (zipCode, zipAdd, zipAddFuri) = Utility.GetAddressFromZipcode(txtZipCode1.Text, txtZipCode2.Text);

            txtZipCode1.Text = zipCode.Substring(0, 3);
            txtZipCode2.Text = zipCode.Substring(3, 4);
            txtAdd.Text = zipAddFuri;
            txtAddKanji.Text = zipAdd;
        }
    }
}
