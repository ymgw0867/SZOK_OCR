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
            shukkoMs(@"C:\Users\kyama\OneDrive - softbit(1)\複合研ディーエル\静岡防犯協会\年別防犯登録票配布シート\31(3603981-.xlsx");
        }

        private void shukkoMs(string sPath)
        {
            cardDataSet dts = new cardDataSet();
            cardDataSetTableAdapters.出庫データTableAdapter adp = new cardDataSetTableAdapters.出庫データTableAdapter();

            Cursor = Cursors.WaitCursor;

            int rNum = 0;

            try
            {
                IXLWorkbook bk;

                int cnt = 0;

                using(bk = new XLWorkbook(sPath, XLEventTracking.Disabled))
                {
                    var sheet1 = bk.Worksheet(1);
                    var tbl = sheet1.RangeUsed().AsTable();

                    //MessageBox.Show(tbl.RowCount().ToString());

                    foreach (var t in tbl.Rows())
                    {
                        //if (t.RowNumber() < 5)
                        //{
                        //    continue;
                        //}

                        if (Utility.nulltoStr2(t.Cell(1).Value) == string.Empty)
                        {
                            continue;
                        }

                        rNum = t.RowNumber();

                        DateTime dt;
                        if (!DateTime.TryParse(t.Cell(2).Value.ToString(), out dt))
                        {
                            //MessageBox.Show(t.RowNumber().ToString() + " " + t.Cell(2).Value.ToString());
                            dt = DateTime.FromOADate(Utility.StrtoDouble(t.Cell(2).Value.ToString()));
                        }

                        // Excelセルデータ取得
                        string num = Utility.nulltoStr2(t.Cell(1).Value).PadLeft(4, '0');
                        int Id = Utility.StrtoInt((dt.Year - 2000).ToString() + dt.Month.ToString("D2") + num);
                        int tNum = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(3).Value));
                        string tName = Utility.nulltoStr2(t.Cell(4).Value);
                        int Busu = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(5).Value));
                        int stNum = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(6).Value));
                        int edNum = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(8).Value));
                        int Uriage = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(10).Value));

                        // 出庫データ追加登録
                        adp.InsertQuery(Id, dt, tNum, tName, Busu, stNum, edNum, Uriage);

                        cnt++;
                    }
                }

                MessageBox.Show("終了しました！  出力件数：" + cnt + "件");
            }
            catch (Exception ex)
            {
                MessageBox.Show(rNum + Environment.NewLine + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
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
