using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SZOK_OCR.Common;

namespace SZOK_OCR.DATA
{
    public partial class frmMakeCsv : Form
    {
        public frmMakeCsv()
        {
            InitializeComponent();
        }
        
        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("静岡県警察本部用CSVファイルを作成しますか？","確認",MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // 静岡県警察本部用CSVファイル作成
            csvOutput();
        }

        /// --------------------------------------------------------------
        /// <summary>
        ///     静岡県警察本部用CSVファイル作成 </summary>
        /// --------------------------------------------------------------
        private void csvOutput()
        {
            // 待機カーソル
            this.Cursor = Cursors.WaitCursor;

            // 防犯登録カードデータ CSVファイル出力クラスインスタンス
            clsOutput p = new clsOutput();

            // 自転車登録.CSVファイル作成
            int c = p.saveCycleCsv();

            // 原付登録.CSVファイル作成
            int a = p.saveAutoCsv();

            // カーソル戻す
            this.Cursor = Cursors.Default;

            // 終了メッセージ表示
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("静岡県警察本部用ＣＳＶデータの出力が終了しました。").Append(Environment.NewLine + Environment.NewLine);
            sb.Append("自転車登録データ：" + c.ToString("#,##0") + "件").Append(Environment.NewLine);
            sb.Append("原付登録データ：" + a.ToString("#,##0") + "件").Append(Environment.NewLine);

            MessageBox.Show(sb.ToString(), "ＣＳＶデータ出力", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void frmMakeCsv_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmCsvReki frm = new frmCsvReki();
            frm.ShowDialog();
            this.Show();
        }
    }
}
