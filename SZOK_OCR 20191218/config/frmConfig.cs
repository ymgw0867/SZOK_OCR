using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SZOK_OCR.Common;
using System.Data.OleDb;

namespace SZOK_OCR.Config
{
    public partial class frmConfig : Form
    {
        public frmConfig()
        {
            InitializeComponent();

            adp.Fill(dts.環境設定);

            var s = dts.環境設定.Single(a => a.ID == global.configKEY);

            if (s.Is受け渡しデータ作成パスNull())
            {
                txtPath2.Text = string.Empty;
            }
            else
            {
                txtPath2.Text = s.受け渡しデータ作成パス;
            }

            if (s.Is郵便番号データパスNull())
            {
                txtPath1.Text = string.Empty;
            }
            else
            {
                txtPath1.Text = s.郵便番号データパス;
            }

            txtDataSpan.Text = s.データ保存月数.ToString();            
        }

        szokDataSetTableAdapters.環境設定TableAdapter adp = new szokDataSetTableAdapters.環境設定TableAdapter();
        szokDataSet dts = new szokDataSet();

        private void frmConfig_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     フォルダダイアログ選択 </summary>
        /// <returns>
        ///     フォルダー名</returns>
        ///------------------------------------------------------------------------
        private string userFolderSelect()
        {
            string fName = string.Empty;

            //出力フォルダの選択ダイアログの表示
            // FolderBrowserDialog の新しいインスタンスを生成する (デザイナから追加している場合は必要ない)
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();

            // ダイアログの説明を設定する
            folderBrowserDialog1.Description = "フォルダを選択してください";

            // ルートになる特殊フォルダを設定する (初期値 SpecialFolder.Desktop)
            folderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.Desktop;

            // 初期選択するパスを設定する
            folderBrowserDialog1.SelectedPath = @"C:\SZOK_OCR";

            // [新しいフォルダ] ボタンを表示する (初期値 true)
            folderBrowserDialog1.ShowNewFolderButton = true;

            // ダイアログを表示し、戻り値が [OK] の場合は、選択したディレクトリを表示する
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                fName = folderBrowserDialog1.SelectedPath + @"\";
            }
            else
            {
                // 不要になった時点で破棄する
                folderBrowserDialog1.Dispose();
                return fName;
            }

            // 不要になった時点で破棄する
            folderBrowserDialog1.Dispose();

            return fName;
        }

        private string userFileSelect()
        {
            DialogResult ret;

            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //ダイアログボックスの初期設定
            openFileDialog1.Title = "郵便番号CSVデータを選択してください";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "CSVファイル(*.CSV)|*.csv|全てのファイル(*.*)|*.*";

            //ダイアログボックスの表示
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel)
            {
                return string.Empty;
            }

            if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " が選択されました。よろしいですか?", "郵便番号CSV確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return string.Empty;
            }

            return openFileDialog1.FileName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 郵便番号CSVデータを選択する
            string sPath = userFileSelect();
            if (sPath != string.Empty)
            {
                txtPath1.Text = sPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // データ更新
            DataUpdate();
        }

        private void DataUpdate()
        {
            // エラーチェック
            if (!errCheck())
            {
                return;
            }

            if (MessageBox.Show("データを更新してよろしいですか","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No) return;
            
            szokDataSet.環境設定Row r = dts.環境設定.Single(a => a.ID == global.configKEY);
            
            r.受け渡しデータ作成パス = txtPath2.Text;
            r.郵便番号データパス = txtPath1.Text;
            //r.データ保存月数 = Utility.StrtoInt(txtDataSpan.Text);   
            r.データ保存月数 = global.flgOff;
            r.更新年月日 = DateTime.Now;

            // データ更新
            adp.Update(r);

            //
            global.cnfPath = r.受け渡しデータ作成パス;
 
            // 終了
            this.Close();
        }

        /// ------------------------------------------------------------------------------------
        /// <summary>
        ///     エラーチェック </summary>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        /// ------------------------------------------------------------------------------------
        private bool errCheck()
        {
            // 郵便番号CSVデータパス
            if (txtPath1.Text.Trim() == string.Empty)
            {
                MessageBox.Show("郵便番号CSVデータパスを入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtPath1.Focus();
                return false;
            }

            // 静岡県警察本部用CSV出力先パス
            if (txtPath2.Text.Trim() == string.Empty)
            {
                MessageBox.Show("静岡県警察本部用CSV出力先フォルダパスを入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtPath2.Focus();
                return false;
            }

            // データ保存月数パス
            if (txtDataSpan.Text.Trim() == string.Empty)
            {
                MessageBox.Show("データ保存月数パスを入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtDataSpan.Focus();
                return false;
            }
            
            return true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmConfig_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //フォルダーを選択する
            string sPath = userFolderSelect();
            if (sPath != string.Empty)
            {
                txtPath2.Text = sPath;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }
    }
}
