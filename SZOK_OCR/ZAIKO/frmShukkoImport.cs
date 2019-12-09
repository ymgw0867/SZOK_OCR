using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SZOK_OCR.Common;
using ClosedXML.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace SZOK_OCR.ZAIKO
{
    public partial class frmShukkoImport : Form
    {
        public frmShukkoImport()
        {
            InitializeComponent();
        }

        cardDataSet dts = new cardDataSet();
        cardDataSetTableAdapters.出庫データTableAdapter adp = new cardDataSetTableAdapters.出庫データTableAdapter();


        private void frmShukkoImport_Load(object sender, EventArgs e)
        {
            label2.Text = "";
            toolStripProgressBar1.Visible = false;

            adp.FillByMaxID(dts.出庫データ);

            var s = dts.出庫データ;

            foreach (var t in s)
            {
                label4.Text = "出庫日：" + t.出庫日.ToShortDateString() + "　店名：" + t.店番 + ":" + t.店名;
            }
        }



        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     Excel日計表から出庫内容を表示する </summary>
        /// <param name="pFile">
        ///     Excel日計表パス</param>
        ///----------------------------------------------------------------------------------
        private void ShukkoMs_Nikkei(string pFile)
        {
            Excel.Application oXls = null;
            Excel.Workbook oXlsBook = null;
            Excel.Worksheet oxlsSheet = null;

            try
            {
                // オブジェクト２次元配列（エクセルシートの内容を受け取る）
                object[,] objArray = null;

                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                oXls = new Excel.Application();

                oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(pFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

                oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range dRng = null;
                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                int cnt = 0;
                int sCnt = 0;
                int rCnt = 0;
                int rNum = 0;
                string msg = "";

                try
                {
                    // 入力管理シートか調べる
                    if (oxlsSheet.Name != "入力管理シート")
                    {
                        throw new Exception("既定の入力管理シートではありません");
                    }

                    // エクセルシートの内容を２次元配列に取得する
                    dRng = oxlsSheet.UsedRange;
                    objArray = dRng.Value2;

                    // 読み込み開始行
                    int fromRow = 5;

                    // 利用領域行数を取得
                    int toRow = oxlsSheet.UsedRange.Rows.Count;

                    // エクセルシートの行を順次読み込む
                    for (int i = fromRow; i <= toRow; i++)
                    {
                        rNum = i;

                        // 店番未記入はネグる
                        if (Utility.nulltoStr2(objArray[i, 3]).Trim() == string.Empty)
                        {
                            continue;
                        }

                        // 日付
                        string sDate = Utility.nulltoStr2(objArray[i, 2]).Trim();

                        DateTime dt;

                        if (!DateTime.TryParse(sDate, out dt))
                        {
                            dt = DateTime.FromOADate(Utility.StrtoDouble(sDate));
                        }

                        // Excelセルデータ取得
                        string num = Utility.nulltoStr2(objArray[i, 1]).Trim().PadLeft(4, '0');
                        int Id = Utility.StrtoInt((dt.Year - 2000).ToString() + dt.Month.ToString("D2") + num);

                        if (adp.FillByID(dts.出庫データ, Id) > 0)
                        {
                            // メッセージ
                            msg = "  登録済みのためスキップしました...... ";
                            sCnt++;
                        }
                        else
                        {
                            int tNum = Utility.StrtoInt(Utility.nulltoStr2(objArray[i, 3]));
                            string tName = Utility.nulltoStr2(objArray[i, 4]);
                            int Busu = Utility.StrtoInt(Utility.nulltoStr2(objArray[i, 5]));
                            int stNum = Utility.StrtoInt(Utility.nulltoStr2(objArray[i, 6]));
                            int edNum = Utility.StrtoInt(Utility.nulltoStr2(objArray[i, 8]));
                            int Uriage = Utility.StrtoInt(Utility.nulltoStr2(objArray[i, 10]));

                            // 出庫データ追加登録
                            adp.InsertQuery(Id, dt, tNum, tName, Busu, stNum, edNum, Uriage);

                            // メッセージ
                            msg = "  登録されました...... ";

                            rCnt++;
                        }

                        cnt++;

                        toolStripProgressBar1.Value = cnt;

                        listBox1.Items.Add(dt.ToShortDateString() + "  " + Utility.nulltoStr2(objArray[i, 4]) + msg + cnt + "/" + toRow);
                        listBox1.TopIndex = listBox1.Items.Count - 1;

                        System.Threading.Thread.Sleep(80);
                        Application.DoEvents();
                    }

                    listBox1.Items.Add("終了しました.....  追加登録：" + rCnt.ToString("#,##0") + "件、登録済スキップ：" + sCnt.ToString("#,##0") + "件");
                    listBox1.TopIndex = listBox1.Items.Count - 1;

                    System.Threading.Thread.Sleep(1000);
                    Application.DoEvents();

                    MessageBox.Show("終了しました", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                catch (Exception ex)
                {
                    MessageBox.Show(rNum + Environment.NewLine + ex.Message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                //Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;
            }
        }


        private void shukkoMs(string sPath)
        {
            Cursor = Cursors.WaitCursor;
            toolStripProgressBar1.Visible = true;

            int rNum = 0;
            string msg = "";

            try
            {
                IXLWorkbook bk;

                int sCnt = 0;
                int rCnt = 0;
                int cnt = 0;

                using (bk = new XLWorkbook(sPath, XLEventTracking.Disabled))
                {
                    var sheet1 = bk.Worksheet(1);
                    var tbl = sheet1.RangeUsed().AsTable();

                    int n = tbl.Rows().Count();

                    toolStripProgressBar1.Minimum = 1;
                    toolStripProgressBar1.Maximum = n;
                    
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

                        if (adp.FillByID(dts.出庫データ, Id) > 0)
                        {
                            // メッセージ
                            msg = "  登録済みのためスキップしました...... ";
                            sCnt++;
                        }
                        else
                        {
                            int tNum = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(3).Value));
                            string tName = Utility.nulltoStr2(t.Cell(4).Value);
                            int Busu = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(5).Value));
                            int stNum = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(6).Value));
                            int edNum = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(8).Value));
                            int Uriage = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(10).Value));

                            // 出庫データ追加登録
                            adp.InsertQuery(Id, dt, tNum, tName, Busu, stNum, edNum, Uriage);

                            // メッセージ
                            msg = "  登録されました...... ";

                            rCnt++;
                        }

                        cnt++;

                        toolStripProgressBar1.Value = cnt;

                        listBox1.Items.Add(dt.ToShortDateString() + "  " + Utility.nulltoStr2(t.Cell(4).Value) + msg + cnt + "/" + n);
                        listBox1.TopIndex = listBox1.Items.Count - 1;

                        System.Threading.Thread.Sleep(80);
                        Application.DoEvents();
                    }
                }

                listBox1.Items.Add("終了しました.....  追加登録：" +  rCnt.ToString("#,##0") + "件、登録済スキップ：" + sCnt.ToString("#,##0") + "件");
                listBox1.TopIndex = listBox1.Items.Count - 1;

                System.Threading.Thread.Sleep(1000);
                Application.DoEvents();

                MessageBox.Show("終了しました","確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "防犯登録カード出庫エクセルデータ選択";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.Filter = "エクセルファイル(*.xlsx,*.xls)|*.xlsx;*.xls|全てのファイル(*.*)|*.*";

            //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
            string fileName;
            DialogResult ret = openFileDialog1.ShowDialog();

            if (ret == System.Windows.Forms.DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                label2.Text = openFileDialog1.FileName;
            }
            else
            {
                fileName = string.Empty;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (label2.Text == string.Empty)
            {
                MessageBox.Show("対象となるExcelファイルを選択してください", "ファイル未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (!System.IO.File.Exists(label2.Text))
            {
                MessageBox.Show("対象となるExcelファイルは存在しません。再度選択してください", "ファイル未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("防犯登録カード出庫データ読み込みを行います。よろしいですか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            button2.Enabled = false;
            button3.Enabled = false;

            // 出庫データ読み込み登録処理
            //shukkoMs(label2.Text);
            ShukkoMs_Nikkei(label2.Text);

            button2.Enabled = true;
            button3.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmShukkoImport_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }
    }
}
