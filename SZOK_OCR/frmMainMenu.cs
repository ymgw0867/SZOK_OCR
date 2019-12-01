using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SZOK_OCR.OCR;
using SZOK_OCR.Common;
using ClosedXML.Excel;

namespace SZOK_OCR
{
    public partial class frmMainMenu : Form
    {
        public frmMainMenu()
        {
            InitializeComponent();
        }

        private void frmMainMenu_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }
        
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmOCR frm = new frmOCR();
            frm.ShowDialog();
            this.Show();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            cardDataSet dtsC = new cardDataSet();
            cardDataSetTableAdapters.SCAN_DATATableAdapter sAdp = new cardDataSetTableAdapters.SCAN_DATATableAdapter();

            szokDataSet dts = new szokDataSet();
            szokDataSetTableAdapters.防犯カードTableAdapter adp = new szokDataSetTableAdapters.防犯カードTableAdapter();

            // 自らのロックファイルを削除する
            Utility.deleteLockFile(Properties.Settings.Default.dataPath, System.Net.Dns.GetHostName());

            //他のPCで処理中の場合、続行不可
            if (Utility.existsLockFile(Properties.Settings.Default.dataPath))
            {
                MessageBox.Show("他のＰＣで処理中です。しばらくおまちください。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // ＯＣＲ認識したスキャンデータ件数
            int s = sAdp.Fill(dtsC.SCAN_DATA);

            // 処理中の防犯登録データ
            int d = adp.Fill(dts.防犯カード);

            // 処理可能なデータが存在するか？
            if (s == 0 && d == 0)
            {
                MessageBox.Show("現在、処理可能なＯＣＲ認識された防犯登録データはありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            //LOCKファイル作成
            Utility.makeLockFile(Properties.Settings.Default.dataPath, System.Net.Dns.GetHostName());

            this.Hide();

            // 処理するデータを取得
            frmFaxSelectHaken frmFax = new frmFaxSelectHaken();
            frmFax.ShowDialog();

            bool _myBool = frmFax.MyBool;
            frmFax.Dispose();

            // ロックファイルを削除する
            Utility.deleteLockFile(Properties.Settings.Default.dataPath, System.Net.Dns.GetHostName());

            if (!_myBool)
            {
                Show();
            }
            else
            {
                // データ作成処理へ
                frmCorrect frm = new frmCorrect(string.Empty);
                frm.ShowDialog();
                Show();
            }











            //this.Hide();
            //frmCorrect frm = new frmCorrect(string.Empty);
            //frm.ShowDialog();
            //this.Show();
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            Config.frmConfig frm = new Config.frmConfig();
            frm.ShowDialog();
            this.Show();
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
            mdbCompact();

        }

        /// ---------------------------------------------------------------------
        /// <summary>
        ///     MDBファイルを最適化する </summary>
        /// ---------------------------------------------------------------------
        private void mdbCompact()
        {
            try
            {
                JRO.JetEngine jro = new JRO.JetEngine();
                string OldDb = Properties.Settings.Default.mdbOlePath;
                string NewDb = Properties.Settings.Default.mdbPathTemp;

                jro.CompactDatabase(OldDb, NewDb);

                //今までのバックアップファイルを削除する
                System.IO.File.Delete(Properties.Settings.Default.mdbPath + global.MDBBACK);

                //今までのファイルをバックアップとする
                System.IO.File.Move(Properties.Settings.Default.mdbPath + global.MDBFILE, Properties.Settings.Default.mdbPath + global.MDBBACK);

                //一時ファイルをMDBファイルとする
                System.IO.File.Move(Properties.Settings.Default.mdbPath + global.MDBTEMP, Properties.Settings.Default.mdbPath + global.MDBFILE);
            }
            catch (Exception e)
            {
                MessageBox.Show("MDB最適化中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
        }
        private void frmMainMenu_Load(object sender, EventArgs e)
        {
            // 環境設定ファイル読み込み
            cnfDataLoad();

            // 防犯カードテーブルに確認フィールドを追加 2016/01/25
            alterSzok();
        }

        private void cnfDataLoad()
        {
            szokDataSet rDts = new szokDataSet();
            szokDataSetTableAdapters.環境設定TableAdapter cAdp = new szokDataSetTableAdapters.環境設定TableAdapter();

            cAdp.Fill(rDts.環境設定);

            // 環境設定ファイル読み込み
            if (rDts.環境設定.Any(a => a.ID == global.configKEY))
            {
                var s = rDts.環境設定.Single(a => a.ID == global.configKEY);
                if (!s.Is受け渡しデータ作成パスNull())
                {
                    global.cnfPath = s.受け渡しデータ作成パス;
                }
                else
                {
                    global.cnfPath = string.Empty;
                }
            }
            else
            {
                MessageBox.Show("環境設定ファイルを登録してください", "システム設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                global.cnfPath = string.Empty;
            }
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            SZOK_OCR.DATA.frmMakeCsv frm = new DATA.frmMakeCsv();
            frm.ShowDialog();
            this.Show();
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            SZOK_OCR.DATA.frmCardList frm = new DATA.frmCardList();
            frm.ShowDialog();
            this.Show();
        }

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     画面チェック確認フィールドを防犯カードテーブルに追加 </summary>
        ///     
        ///     2016/05/25 : 車両番号１のサイズ1から2に変更
        ///--------------------------------------------------------------------------
        private void alterSzok()
        {
            OleDbConnection cn = new OleDbConnection();
            //cn.ConnectionString = Properties.Settings.Default.SZOKConnectionString;
            //cn.Open();

            OleDbCommand sCom = new OleDbCommand();

            //try
            //{
            //    sCom.Connection = cn;
            //    //sCom.CommandText = "ALTER TABLE 防犯カード ADD COLUMN 確認 int default 0 ";
            //    sCom.CommandText = "ALTER TABLE 防犯カード ALTER COLUMN 車両番号1 TEXT(2)";
            //    sCom.ExecuteNonQuery();
            //}
            //catch (Exception)
            //{
            //    // 何もしない
            //}
            //finally
            //{
            //    if (sCom.Connection.State == ConnectionState.Open)
            //    {
            //        sCom.Connection.Close();
            //    }
            //}

            // サーバー@防犯登録データテーブル
            cn.ConnectionString = Properties.Settings.Default.SZOK_CARDConnectionString;
            cn.Open();

            try
            {
                sCom.Connection = cn;
                //sCom.CommandText = "ALTER TABLE 防犯登録データ ALTER COLUMN 車両番号1 TEXT(2)";
                sCom.CommandText = "ALTER TABLE 防犯登録データ ADD COLUMN 除外 int default 0 ";
                sCom.ExecuteNonQuery();
            }
            catch (Exception)
            {
                // 何もしない
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }
            }

        }

        private void LinkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            SZOK_OCR.DATA.frmScanList frm = new DATA.frmScanList();
            frm.ShowDialog();
            this.Show();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            // 移動先フォルダがあるか？なければ作成する（TIFフォルダ内の年月フォルダ）
            if (!System.IO.Directory.Exists(@"\\YUDAI-PC\Users\kyama\test"))
            {
                System.IO.Directory.CreateDirectory(@"\\YUDAI-PC\Users\kyama\test");
            }

            string f1 = @"C:\SZOK_OCR\f1\CCF20160106.tif";
            //string f2 = @"C:\SZOK_OCR\f2\CCF20160106.tif";
            string f2 = @"\\YUDAI-PC\Users\kyama\test";

            // ファイルを移動する
            if (System.IO.File.Exists(f1))
            {
                System.IO.File.Move(f1, f2);
            }

            MessageBox.Show("fin!");
        }

        private void linkLabel8_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            SZOK_OCR.ZAIKO.frmZaikoMenu frm = new ZAIKO.frmZaikoMenu();
            frm.ShowDialog();
            this.Show();
        }
    }
}
