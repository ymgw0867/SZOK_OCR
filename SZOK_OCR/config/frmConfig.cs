using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SZOK_OCR.Common;

namespace SZOK_OCR.Config
{
    public partial class frmConfig : Form
    {
        public frmConfig()
        {
            InitializeComponent();

            // コメント化：2026/08/17
            //adp.Fill(dts.環境設定);

            //var s = dts.環境設定.Single(a => a.ID == global.configKEY);

            //if (s.Is受け渡しデータ作成パスNull())
            //{
            //    txtPath2.Text = string.Empty;
            //}
            //else
            //{
            //    txtPath2.Text = s.受け渡しデータ作成パス;
            //}

            //if (s.Is郵便番号データパスNull())
            //{
            //    txtPath1.Text = string.Empty;
            //}
            //else
            //{
            //    txtPath1.Text = s.郵便番号データパス;
            //}

            //txtDataSpan.Text = s.データ保存月数.ToString();
            //

            txtPath2.Text = string.Empty;

            // 環境設定データを取得
            master = new ClsMaster(Properties.Settings.Default.sServerName, Properties.Settings.Default.sLogin,
                                   Properties.Settings.Default.sPass, Properties.Settings.Default.sDatabase);
            //conn = master.OpenConnection();
            configData = master.GetData<TblConfig>(global.configKEY.ToString());

            if (configData != null)
            {
                txtPath1.Text = configData.ZipCodePath ?? string.Empty;
                txtDataSpan.Text = configData.DataSaveMonth.ToString();
            }
            else
            {
                // データが存在しない場合は、テキストボックスを空にする
                txtPath1.Text = string.Empty;
                txtDataSpan.Text = string.Empty;

                // メッセージ表示
                MessageBox.Show("環境設定データが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // コメント化：2026/08/17
        //szokDataSetTableAdapters.環境設定TableAdapter adp = new szokDataSetTableAdapters.環境設定TableAdapter();
        //szokDataSet dts = new szokDataSet();

        // SQL接続
        SqlConnection conn;

        // マスタークラス
        ClsMaster master;

        // 環境設定データ
        TblConfig configData;

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

            // コメント化：2026/08/17
            //szokDataSet.環境設定Row r = dts.環境設定.Single(a => a.ID == global.configKEY);

            //r.受け渡しデータ作成パス = txtPath2.Text;
            //r.郵便番号データパス = txtPath1.Text;
            //r.データ保存月数 = global.flgOff;
            //r.更新年月日 = DateTime.Now;

            //// データ更新
            //adp.Update(r);

            // 環境設定データを更新
            configData.ZipCodePath = txtPath1.Text;
            configData.DataSaveMonth = Utility.StrtoInt(txtDataSpan.Text);
            configData.UpDate = DateTime.Now;

            master.UpDate(configData);

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

            // コメント化：2026/08/17
            //// 静岡県警察本部用CSV出力先パス
            //if (txtPath2.Text.Trim() == string.Empty)
            //{
            //    MessageBox.Show("静岡県警察本部用CSV出力先フォルダパスを入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    txtPath2.Focus();
            //    return false;
            //}

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
            //// SQL接続を閉じる
            //master.CloseConnection(conn);

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
