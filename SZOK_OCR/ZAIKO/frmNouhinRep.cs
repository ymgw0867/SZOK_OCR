using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel;
using SZOK_OCR.Common;
using Excel = Microsoft.Office.Interop.Excel;

namespace SZOK_OCR.ZAIKO
{
    public partial class frmNouhinRep : Form
    {
        public frmNouhinRep()
        {
            InitializeComponent();
        }

        // データグリッドビューカラム定義
        string colUCode = "col1";
        string colUName = "col2";
        string colDate = "col3";
        string colShukko = "col4";
        string colSNum = "col7";
        string colENum = "col8";
        string colDaikin = "col9";

        private void button3_Click(object sender, EventArgs e)
        {
            // 日計表選択
            if (GetExcelSheet())
            {
                // 画面表示
                SearchData(label1.Text, dataGridView1);
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///--------------------------------------------------------------------
        private void gridViewSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ ゴシック", 10, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", 10, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 362;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                //各列幅指定
                tempDGV.Columns.Add(colDate, "出庫日");
                tempDGV.Columns.Add(colUCode, "コード");
                tempDGV.Columns.Add(colUName, "得意先名");
                tempDGV.Columns.Add(colShukko, "出庫部数");
                tempDGV.Columns.Add(colSNum, "開始番号");
                tempDGV.Columns.Add(colENum, "終了番号");
                tempDGV.Columns.Add(colDaikin, "代金");

                tempDGV.Columns[colDate].Width = 110;
                tempDGV.Columns[colUCode].Width = 90;
                tempDGV.Columns[colUName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[colShukko].Width = 80;
                tempDGV.Columns[colSNum].Width = 110;
                tempDGV.Columns[colENum].Width = 110;
                tempDGV.Columns[colDaikin].Width = 110;

                tempDGV.Columns[colUCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colShukko].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colSNum].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colENum].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colDaikin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                tempDGV.ReadOnly = true;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更可
                tempDGV.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                tempDGV.StandardTab = true;

                // 罫線
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowExcelData(string sPath, DataGridView g)
        {
            Cursor = Cursors.WaitCursor;

            int rNum = 0;
            string msg = "";

            try
            {
                IXLWorkbook bk;

                int cnt = 0;

                using (bk = new XLWorkbook(sPath, XLEventTracking.Disabled))
                {
                    var sheet1 = bk.Worksheet("入力管理シート");
                    var tbl = sheet1.RangeUsed().AsTable();

                    int n = tbl.Rows().Count();

                    foreach (var t in tbl.Rows())
                    {
                        if (t.RowNumber() < 5)
                        {
                            continue;
                        }

                        // 店番
                        if (Utility.nulltoStr2(t.Cell(3).Value) == string.Empty)
                        {
                            continue;
                        }

                        rNum = t.RowNumber();

                        DateTime dt;

                        if (!DateTime.TryParse(t.Cell(2).Value.ToString(), out dt))
                        {
                            dt = DateTime.FromOADate(Utility.StrtoDouble(t.Cell(2).Value.ToString()));
                        }

                        // Excelセルデータ取得
                        string num = Utility.nulltoStr2(t.Cell(1).Value).PadLeft(4, '0');
                        int Id = Utility.StrtoInt((dt.Year - 2000).ToString() + dt.Month.ToString("D2") + num);

                        g.Rows.Add();
                        g[colDate, cnt].Value = dt.ToShortDateString();
                        g[colUCode, cnt].Value = Utility.nulltoStr2(t.Cell(3).Value);
                        g[colUName, cnt].Value = Utility.nulltoStr2(t.Cell(4).Value);
                        g[colShukko, cnt].Value = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(5).Value));
                        g[colSNum, cnt].Value = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(6).Value));
                        g[colENum, cnt].Value = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(8).Value));
                        g[colDaikin, cnt].Value = Utility.StrtoInt(Utility.nulltoStr2(t.Cell(10).Value));

                        cnt++;
                    }
                }
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



        private bool GetExcelSheet()
        {
            bool rtn = false;

            openFileDialog1.Title = "日計表エクセルファイル選択";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.Filter = "エクセルファイル(*.xlsx,*.xls)|*.xlsx;*.xls|全てのファイル(*.*)|*.*";

            //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
            string fileName;
            DialogResult ret = openFileDialog1.ShowDialog();

            if (ret == System.Windows.Forms.DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                label1.Text = openFileDialog1.FileName;
                rtn = true;
            }
            else
            {
                fileName = string.Empty;
            }

            return rtn;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmNouhinRep_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void frmNouhinRep_Load(object sender, EventArgs e)
        {
            button1.Enabled = false;

            gridViewSetting(dataGridView1);
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     Excel日計表から出庫内容を表示する </summary>
        /// <param name="pFile">
        ///     Excel日計表パス</param>
        ///----------------------------------------------------------------------------------
        private void SearchData(string pFile, DataGridView g)
        {
            // オブジェクト２次元配列（エクセルシートの内容を受け取る）
            object[,] objArray = null;

            //マウスポインタを待機にする
            this.Cursor = Cursors.WaitCursor;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(pFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                           Type.Missing, Type.Missing));
           
            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

            Excel.Range dRng = null;
            Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

            int dCnt = 0;
            int cnt = 0;

            try
            {
                // 入力管理シートか調べる
                if (oxlsSheet.Name != "入力管理シート")
                {
                    throw new Exception("既定の入力管理シートではありません");
                }

                // エクセルシートの内容を２次元配列に取得する
                //dRng = oxlsSheet.Range[oxlsSheet.Cells[6, 23], oxlsSheet.Cells[55, 28]];
                dRng = oxlsSheet.UsedRange;
                objArray = dRng.Value2;

                // 読み込み開始行
                int fromRow = 5;

                // 利用領域行数を取得
                int toRow = oxlsSheet.UsedRange.Rows.Count;

                // エクセルシートの行を順次読み込む
                for (int i = fromRow; i <= toRow; i++)
                {
                    // 店番未記入はネグる
                    if (Utility.nulltoStr2(objArray[i, 3]).Trim() == string.Empty)
                    {
                        continue;
                    }

                    g.Rows.Add();

                    // 日付
                    string sDate = Utility.nulltoStr2(objArray[i, 2]).Trim();

                    DateTime dt;

                    if (!DateTime.TryParse(sDate, out dt))
                    {
                        dt = DateTime.FromOADate(Utility.StrtoDouble(sDate));
                    }

                    g[colDate, cnt].Value = dt.ToShortDateString();

                    // 店番
                    g[colUCode, cnt].Value = Utility.nulltoStr2(objArray[i, 3]).Trim();

                    // 店名
                    g[colUName, cnt].Value = Utility.nulltoStr2(objArray[i, 4]).Trim();

                    // 出庫部数
                    g[colShukko, cnt].Value = Utility.nulltoStr2(objArray[i, 5]).Trim();

                    // 開始CPA番号
                    g[colSNum, cnt].Value = Utility.nulltoStr2(objArray[i, 6]).Trim();

                    // 終了CPA番号
                    g[colENum, cnt].Value = Utility.nulltoStr2(objArray[i, 8]).Trim();

                    // 代金
                    g[colDaikin, cnt].Value = Utility.nulltoStr2(objArray[i, 10]).Trim();

                    cnt++;
                }

                if (cnt > 0)
                {
                    g.CurrentCell = null;
                }
            }

            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エクセルシートオープンエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            GetGridData(dataGridView1, e.RowIndex);
        }

        private void GetGridData(DataGridView g, int r)
        {
            txtTenName.Text = Utility.nulltoStr2(g[colUName, r].Value);
            txtBusu.Text = Utility.nulltoStr2(g[colShukko, r].Value);
            txtSNum.Text = Utility.nulltoStr2(g[colSNum, r].Value);
            txtENum.Text = Utility.nulltoStr2(g[colENum, r].Value);
            txtDaikin.Text = Utility.nulltoStr2(g[colDaikin, r].Value);
        }

        private void txtTenName_TextChanged(object sender, EventArgs e)
        {
            if (txtTenName.Text.Trim() != string.Empty)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     防犯登録カード納品書・請求書印刷 </summary>
        ///---------------------------------------------------------------
        private void NouhinReport()
        {
            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsNouhinRep, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    // 納品書・請求書
                    oxlsSheet.Cells[1, 9] = "伝票№ " + txtDenNum.Text;
                    oxlsSheet.Cells[2, 5] = DateTime.Today.ToLongDateString();
                    oxlsSheet.Cells[5, 1] = txtTenName.Text;
                    oxlsSheet.Cells[12, 1] = "　（ＣＰＡ　" + txtSNum.Text + " ～ " + txtENum.Text + "）";
                    oxlsSheet.Cells[12, 5] = (Utility.StrtoInt(txtBusu.Text) / 10) + "セット";
                    oxlsSheet.Cells[12, 9] = Utility.StrtoInt(txtDaikin.Text);
                    oxlsSheet.Cells[17, 9] = Utility.StrtoInt(txtDaikin.Text);

                    // 控
                    oxlsSheet.Cells[25, 9] = "伝票№ " + txtDenNum.Text;
                    oxlsSheet.Cells[26, 5] = DateTime.Today.ToLongDateString();
                    oxlsSheet.Cells[29, 1] = txtTenName.Text;
                    oxlsSheet.Cells[36, 1] = "　（ＣＰＡ　" + txtSNum.Text + " ～ " + txtENum.Text + "）";
                    oxlsSheet.Cells[36, 5] = (Utility.StrtoInt(txtBusu.Text) / 10) + "セット";
                    oxlsSheet.Cells[36, 9] = Utility.StrtoInt(txtDaikin.Text);
                    oxlsSheet.Cells[41, 9] = Utility.StrtoInt(txtDaikin.Text);

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    oXls.Visible = true;

                    //印刷
                    oxlsSheet.PrintPreview(true);

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "納品書・請求書印刷エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "納品書・請求書印刷エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            NouhinReport();
        }
    }
}
