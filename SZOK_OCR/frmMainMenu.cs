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
            this.Hide();
            frmCorrect frm = new frmCorrect(string.Empty);
            frm.ShowDialog();
            this.Show();
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
    }
}
