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

            button2.Enabled = false;  // 変更届登録ボタンは初期状態では無効化
            button3.Enabled = true;   // 閉じるボタンは有効化

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
            lblZipCode1.Text = r.ZipCode1;
            lblZipCode2.Text = r.ZipCode2;
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

            if (IsCheckBoxChecked())
            {
                button2.Enabled = true;  // 変更届登録ボタンを有効化
            }
            else
            {
                button2.Enabled = false; // 変更届登録ボタンを無効化
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

            if (IsCheckBoxChecked())
            {
                button2.Enabled = true;  // 変更届登録ボタンを有効化
            }
            else
            {
                button2.Enabled = false; // 変更届登録ボタンを無効化
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

            if (IsCheckBoxChecked())
            {
                button2.Enabled = true;  // 変更届登録ボタンを有効化
            }
            else
            {
                button2.Enabled = false; // 変更届登録ボタンを無効化
            }
        }

        /// <summary>
        /// 何かしらの変更チェックボックスが選択されているかを確認する
        /// </summary>
        /// <returns>true: 何かしらの変更チェックボックスが選択されている, false: 何も選択されていない</returns>
        private bool IsCheckBoxChecked()
        {
            return ChkName.Checked || ChkAddress.Checked || ChkTel.Checked;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 何かしらの変更チェックボックスが選択されているかをチェック
            if (!IsCheckBox())
            {
                return;
            }

            // 氏名のチェックボックスが選択されている場合、変更後氏名入力されているかを確認する
            if (!IsName())
            {
                return;
            }

            // 住所のチェックボックスが選択されている場合、郵便番号と住所が入力されているかを確認する
            if (!IsAddress())
            {
                return;
            }

            // 電話番号のチェックボックスが選択されている場合、電話番号が入力されているかを確認する
            if (!IsTelephone())
            {
                return;
            }

            if (MessageBox.Show($"変更届を登録します。{Environment.NewLine}変更日付：{dateTimePicker1.Value.ToShortDateString()}{Environment.NewLine}よろしいですか？", "変更届登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // SQL Server接続
            var master = new ClsMaster(Properties.Settings.Default.sServerName, Properties.Settings.Default.sLogin,
                                   Properties.Settings.Default.sPass, Properties.Settings.Default.sDatabase);

            // 変更届登録・防犯登録データ更新
            master.ChangeNotification(new bool[] { ChkName.Checked, ChkAddress.Checked, ChkTel.Checked }, _ix, ChangeNotification());

            // 変更届登録完了ステータス
            EditMode = true;

            // 変更届登録完了メッセージを表示してフォームを閉じる
            MessageBox.Show("変更届を登録しました", "変更届登録完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
        }

        /// <summary>
        /// 変更届の内容をTblChangeNotificationに格納する
        /// </summary>
        private TblChangeNotification ChangeNotification()
        {
            bool[] changeFlags = new bool[3]; // 変更フラグ配列: [0]氏名, [1]住所, [2]電話番号
            changeFlags[0] = ChkName.Checked;
            changeFlags[1] = ChkAddress.Checked;
            changeFlags[2] = ChkTel.Checked;

            var changeData = new TblChangeNotification()
            {
                UpdateDay = DateTime.Parse(dateTimePicker1.Value.ToShortDateString()),
                Number = lblNumber.Text,

                OldName = changeFlags[0] ? lblName.Text : "",
                NewName = changeFlags[0] ? txtName.Text : "",       

                OldZipCode1 = changeFlags[1] ? lblZipCode1.Text: "",
                OldZipCode2 = changeFlags[1] ? lblZipCode2.Text : "",
                NewZipCode1 = changeFlags[1] ? txtZipCode1.Text : "",
                NewZipCode2 = changeFlags[1] ? txtZipCode2.Text : "",
                OldAddressKanji = changeFlags[1] ? lblAddKanji.Text : "",
                OldAddress1 = changeFlags[1] ? lblAddress.Text : "",
                OldAddress2 = "", // 旧住所2は未使用
                NewAddressKanji = changeFlags[1] ? txtAddKanji.Text : "",
                NewAddress1 = changeFlags[1] ? txtAdd.Text : "",
                NewAddress2 = "", // 新住所2は未使用

                OldMobile1 = changeFlags[2] ? lblTel1.Text : "",
                OldMobile2 = changeFlags[2] ? lblTel2.Text : "",
                OldMobile3 = changeFlags[2] ? lblTel3.Text : "",
                NewMobile1 = changeFlags[2] ? txtTel1.Text : "",
                NewMobile2 = changeFlags[2] ? txtTel2.Text : "",
                NewMobile3 = changeFlags[2] ? txtTel3.Text : "",
                Memo = "", // メモは未使用
                UpDate = DateTime.Now
            };
            return changeData;
        }

        /// <summary>
        /// チェックボックスが1つも選択されていない場合、メッセージを表示してfalseを返す
        /// </summary>
        /// <returns>true: チェックボックスが1つ以上選択されている, false: チェックボックスが1つも選択されていない</returns>
        private bool IsCheckBox()
        {
            if (ChkName.Checked == false && ChkAddress.Checked == false && ChkTel.Checked == false)
            {
                MessageBox.Show("変更する項目を選択してください", "変更届", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 氏名のチェックボックスが選択されている場合、氏名が入力されているかを確認する。入力されていない場合、メッセージを表示してfalseを返す
        /// </summary>
        /// <returns>true: 氏名が入力されている, false: 氏名が入力されていない</returns>
        private bool IsName()
        {
            if (ChkName.Checked && txtName.Text.Trim() == string.Empty)
            {
                MessageBox.Show("変更後の氏名を入力してください", "変更届", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 住所のチェックボックスが選択されている場合、郵便番号と住所が入力されているかを確認する。入力されていない場合、メッセージを表示してfalseを返す
        /// </summary>
        /// <returns>true: 住所が入力されている, false: 住所が入力されていない</returns>
        private bool IsAddress()
        {
            if (ChkAddress.Checked && (txtZipCode1.Text.Trim() == string.Empty || txtZipCode2.Text.Trim() == string.Empty || txtAdd.Text.Trim() == string.Empty || txtAddKanji.Text.Trim() == string.Empty))
            {
                MessageBox.Show("変更後の住所を入力してください", "変更届", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 電話番号のチェックボックスが選択されている場合、電話番号が入力されているかを確認する。入力されていない場合、メッセージを表示してfalseを返す
        /// </summary>
        /// <returns>true: 電話番号が入力されている, false: 電話番号が入力されていない</returns>
        private bool IsTelephone()
        {
            if (ChkTel.Checked && (txtTel1.Text.Trim() == string.Empty || txtTel2.Text.Trim() == string.Empty || txtTel3.Text.Trim() == string.Empty))
            {
                MessageBox.Show("変更後の電話番号を入力してください", "変更届", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
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

        private void txtZipCode1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        private void txtName_Leave(object sender, EventArgs e)
        {
            TextBox txtbox = (TextBox)sender;
            txtbox.Text = Utility.getStrConv(txtbox.Text);
        }
    }
}
