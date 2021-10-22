using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SZOK_OCR.Common;

namespace SZOK_OCR.ZAIKO
{
    public partial class frmKaishuData : Form
    {
        public frmKaishuData()
        {
            InitializeComponent();
        }

        cardDataSet dts = new cardDataSet();
        cardDataSetTableAdapters.SCAN_DATATableAdapter sAdp = new cardDataSetTableAdapters.SCAN_DATATableAdapter();
        cardDataSetTableAdapters.防犯登録データTableAdapter dAdp = new cardDataSetTableAdapters.防犯登録データTableAdapter();
        cardDataSetTableAdapters.出庫データTableAdapter shuAdp = new cardDataSetTableAdapters.出庫データTableAdapter();
        cardDataSetTableAdapters.回収データTableAdapter kaiAdp = new cardDataSetTableAdapters.回収データTableAdapter();
        
        private void frmKaishuData_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     SCAN_DATA方から回収データを登録する </summary>
        ///-----------------------------------------------------------------
        private void SDataToKaishu()
        {
            DateTime dtn = DateTime.Now;
            listBox1.Items.Add("処理を開始しました..... " + dtn.ToShortDateString() + "   " + dtn.ToShortTimeString() + ":" + dtn.Second);

            int cnt = 0;
            int rCnt = 0;

            //int n = sAdp.FillByNonKaishu(dts.SCAN_DATA);      // 20200720 コメント化
            int n = sAdp.FillByNonKaishu202007(dts.SCAN_DATA);  // 20200720 登録番号重複を防ぐため登録番号で照合

            shuAdp.Fill(dts.出庫データ);

            toolStripProgressBar1.Minimum = 1;
            toolStripProgressBar1.Maximum = n;

            Cursor = Cursors.WaitCursor;

            try
            {
                foreach (var t in dts.SCAN_DATA.OrderBy(a => a.ID))
                {
                    cnt++;

                    string ocrdt = t.画像名.Substring(0, 4) + "/" + t.画像名.Substring(4, 2) + "/" + t.画像名.Substring(6, 2);

                    int Num = Utility.StrtoInt(t.登録番号.Replace("CPA", "").Trim());

                    foreach (var s in dts.出庫データ.Where(a => a.開始登録番号 <= Num && a.終了登録番号 >= Num))
                    {
                        DateTime dt;

                        if (DateTime.TryParse(ocrdt, out dt))
                        {
                            // 回収データ登録
                            kaiAdp.InsertQuery(s.ID, dt, Num, global.flgOff, t.ID, DateTime.Now);

                            rCnt++;

                            toolStripProgressBar1.Value = cnt;
                            listBox1.Items.Add(ocrdt + " " + t.登録番号 + " " + s.店名 + ".....  " + cnt + "/" + n);
                            listBox1.TopIndex = listBox1.Items.Count - 1;

                            System.Threading.Thread.Sleep(100);
                            Application.DoEvents();
                        }
                    }
                }

                dtn = DateTime.Now;
                listBox1.Items.Add("終了しました.....  登録：" + rCnt.ToString("#,##0") + "件    " + 
                    dtn.ToShortDateString() + " " + dtn.ToShortTimeString() + ":" + dtn.Second);
                listBox1.TopIndex = listBox1.Items.Count - 1;

                System.Threading.Thread.Sleep(1000);
                Application.DoEvents();

                MessageBox.Show("終了しました", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(cnt + " " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     SCAN_DATA方から回収データを登録する </summary>
        ///-----------------------------------------------------------------
        private void BDataToKaishu()
        {
            DateTime dtn = DateTime.Now;
            listBox1.Items.Add("処理を開始しました..... " + dtn.ToShortDateString() + "   " + dtn.ToShortTimeString() + ":" + dtn.Second);

            int cnt = 0;
            int rCnt = 0;

            //int n = dAdp.FillByNonKaishu(dts.防犯登録データ);        // 2020/07/20 コメント化
            int n = dAdp.FillByNonkaishu202007(dts.防犯登録データ);    // 登録番号重複を防ぐために登録番号で照合 2020/07/20

            shuAdp.Fill(dts.出庫データ);

            toolStripProgressBar1.Minimum = 1;
            toolStripProgressBar1.Maximum = n;

            Cursor = Cursors.WaitCursor;

            try
            {
                foreach (var t in dts.防犯登録データ.OrderBy(a => a.ID))
                {
                    cnt++;

                    string ocrdt = t.画像名.Substring(0, 4) + "/" + t.画像名.Substring(4, 2) + "/" + t.画像名.Substring(6, 2);

                    int Num = Utility.StrtoInt(t.登録番号.Replace("CPA", "").Trim());

                    foreach (var s in dts.出庫データ.Where(a => a.開始登録番号 <= Num && a.終了登録番号 >= Num))
                    {
                        DateTime dt;

                        if (DateTime.TryParse(ocrdt, out dt))
                        {
                            // 回収データ登録
                            kaiAdp.InsertQuery(s.ID, dt, Num, t.ID, global.flgOff, DateTime.Now);

                            rCnt++;

                            toolStripProgressBar1.Value = cnt;
                            listBox1.Items.Add(ocrdt + " " + t.登録番号 + " " + s.店名 + ".....  " + cnt + "/" + n);
                            listBox1.TopIndex = listBox1.Items.Count - 1;

                            System.Threading.Thread.Sleep(100);
                            Application.DoEvents();
                        }
                    }
                }

                dtn = DateTime.Now;
                listBox1.Items.Add("終了しました.....  登録：" + rCnt.ToString("#,##0") + "件    " +
                    dtn.ToShortDateString() + " " + dtn.ToShortTimeString() + ":" + dtn.Second);
                listBox1.TopIndex = listBox1.Items.Count - 1;

                System.Threading.Thread.Sleep(1000);
                Application.DoEvents();

                MessageBox.Show("終了しました", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("防犯登録カード回収処理を行います。よろしいですか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            if (comboBox1.SelectedIndex == 0)
            {
                // SCAN_DATAから回収データを作成
                SDataToKaishu();
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                // 防犯登録データから回収データを作成（過去データ）
                BDataToKaishu();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmKaishuData_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }
    }
}
