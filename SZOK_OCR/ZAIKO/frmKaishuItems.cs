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

namespace SZOK_OCR.ZAIKO
{
    public partial class frmKaishuItems : Form
    {
        public frmKaishuItems(ClsShukko shukko)
        {
            InitializeComponent();

            clsShukko = shukko; 
        }

        Image OcrImg = null;

        // 画像サイズ
        float B_WIDTH  = 0.433f;
        float B_HEIGHT = 0.433f;
        //float B_WIDTH = 0.345f;
        //float B_HEIGHT = 0.345f;
        float n_width = 0f;
        float n_height = 0f;

        cardDataSet dts = new cardDataSet();
        cardDataSetTableAdapters.回収データTableAdapter kAdp = new cardDataSetTableAdapters.回収データTableAdapter();
        cardDataSetTableAdapters.防犯登録データTableAdapter tAdp = new cardDataSetTableAdapters.防犯登録データTableAdapter();
        cardDataSetTableAdapters.SCAN_DATATableAdapter sAdp = new cardDataSetTableAdapters.SCAN_DATATableAdapter();

        ClsShukko clsShukko = new ClsShukko();

        // カラム定義
        private readonly string colSeq       = "c0";
        private readonly string colNum       = "c1";
        private readonly string colStatus    = "c2";
        private readonly string colDate      = "c3";
        private readonly string colID        = "c4";
        private readonly string colTourokuID = "c5";
        private readonly string colScanID    = "c6";

        private void frmKaishuItems_Load(object sender, EventArgs e)
        {
            lblUCode.Text      = clsShukko.UCode.ToString();
            lblUName.Text      = clsShukko.User;
            lblShukkoDate.Text = clsShukko.ShukkoDate;
            lblStartNum.Text   = clsShukko.StartNumber.ToString();
            lblEndNum.Text     = clsShukko.EndNumber.ToString();
            lblSuu.Text        = clsShukko.Suu.ToString();
            lblKaishu.Text     = clsShukko.Kaishu.ToString();
            lblZan.Text        = clsShukko.Zan.ToString();


            // DataGridView定義
            GridviewSet(dg1);

            // 出庫IDをキーとしてで回収データを取得
            DateTime dt3 = new DateTime(clsShukko.kaishuLimitDate.Year, clsShukko.kaishuLimitDate.Month, clsShukko.kaishuLimitDate.Day, 23, 59, 59); // 2021/10/22
            kAdp.FillByKaishu(dts.回収データ, clsShukko.ID, dt3);

            // 回収データ明細表示
            KaishuDataShow();
        }

        private void KaishuDataShow()
        {
            // 行数を設定して表示色を初期化
            dg1.Rows.Clear();
            dg1.Rows.Add(clsShukko.Suu);
            
            int seqNum = 1;
            for (int i = clsShukko.StartNumber; i <= clsShukko.EndNumber; i++)
            {
                dg1[colSeq,       seqNum - 1].Value = seqNum;
                dg1[colNum,       seqNum - 1].Value = i;

                dg1[colStatus,    seqNum - 1].Value = "";
                dg1[colDate,      seqNum - 1].Value = "";
                dg1[colID,        seqNum - 1].Value = "";
                dg1[colTourokuID, seqNum - 1].Value = global.FLGOFF;
                dg1[colScanID,    seqNum - 1].Value = global.FLGOFF;

                foreach (var t in dts.回収データ.Where(a => a.登録番号 == i))
                {
                    dg1[colStatus, seqNum - 1].Value    = "○";
                    dg1[colDate,   seqNum - 1].Value    = t.回収年月日.ToShortDateString();
                    dg1[colID,     seqNum - 1].Value    = t.ID;
                    dg1[colTourokuID, seqNum - 1].Value = t.防犯登録ID;
                    dg1[colScanID, seqNum - 1].Value    = t.SCANID;
                }

                // 未回収行背景色
                if (Utility.nulltoStr2(dg1[colID, seqNum - 1].Value) == "")
                {
                    dg1.Rows[seqNum - 1].DefaultCellStyle.BackColor = Color.MistyRose;
                }

                //カレントセル選択状態としない
                dg1.CurrentCell = null;

                seqNum++;
            }
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     出勤簿データグリッドビュー定義 </summary>
        ///------------------------------------------------------------------------
        private void GridviewSet(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する
                tempDGV.EnableHeadersVisualStyles = false;
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                tempDGV.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                //tempDGV.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ ゴシック", (float)(9), FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", 10, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 22;
                tempDGV.RowTemplate.Height  = 22;

                // 全体の高さ
                tempDGV.Height = 684;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                // 列追加
                tempDGV.Columns.Add(colSeq,       "");
                tempDGV.Columns.Add(colNum,       "登録番号");
                tempDGV.Columns.Add(colStatus,    "回収");
                tempDGV.Columns.Add(colDate,      "回収日");
                tempDGV.Columns.Add(colID,        "");
                tempDGV.Columns.Add(colTourokuID, "");
                tempDGV.Columns.Add(colScanID,    "");

                // 各列幅指定
                tempDGV.Columns[colSeq].Width    = 50;
                tempDGV.Columns[colNum].Width    = 100;
                tempDGV.Columns[colStatus].Width = 110;
                //tempDGV.Columns[colDate].Width   = 165;

                tempDGV.Columns[colID].Visible        = false;
                tempDGV.Columns[colTourokuID].Visible = false;
                tempDGV.Columns[colScanID].Visible    = false;

                tempDGV.Columns[colDate].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[colSeq].DefaultCellStyle.Alignment    = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colNum].DefaultCellStyle.Alignment    = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colStatus].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colDate].DefaultCellStyle.Alignment   = DataGridViewContentAlignment.MiddleCenter;

                // 編集可否
                tempDGV.ReadOnly = true;

                // 列ごとの設定
                //foreach (DataGridViewColumn c in tempDGV.Columns)
                //{
                //    // 編集可否
                //    if (c.Name == colNum || c.Name == colHinName)
                //    {
                //        c.ReadOnly = true;
                //    }
                //    else
                //    {
                //        c.ReadOnly = false;
                //    }

                //    // フォントサイズ
                //    if (c.Name == colSuu)
                //    {
                //        c.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", 12, FontStyle.Regular);
                //    }
                //}

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect   = false;

                // 編集モード
                //tempDGV.EditMode = DataGridViewEditMode.EditOnEnter;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                tempDGV.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                //TAB動作
                tempDGV.StandardTab = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
                //tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // ソート不可
                //foreach (DataGridViewColumn c in tempDGV.Columns)
                //{
                //    c.SortMode = DataGridViewColumnSortMode.NotSortable;
                //}

                // 罫線
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
                //tempDGV.GridColor = Color.SteelBlue;

                // コンテキストメニュー
                //tempDGV.ContextMenuStrip = this.contextMenuStrip1;

                // 文字色
                //tempDGV.RowsDefaultCellStyle.ForeColor = Global.defaultColor;
                //tempDGV.RowsDefaultCellStyle.ForeColor = Color.Black;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmKaishuItems_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void frmKaishuItems_Shown(object sender, EventArgs e)
        {
            dg1.CurrentCell = null;

            pictureBox1.Image = null;
        }

        private void ShowImage(string imgPath, string img)
        {
            // 画像表示
            string _img = imgPath + img;

            if (System.IO.File.Exists(_img))
            {
                // System.Drawing.Imageを作成する
                OcrImg = Utility.CreateImage(_img);
                ImgShow(OcrImg, B_WIDTH, B_HEIGHT);
                //trackBar1.Enabled = true;
                //btnLeft.Enabled = true;
            }
            else
            {
                pictureBox1.Image = null;
                //trackBar1.Enabled = false;
                //btnLeft.Enabled = false;
            }
        }

        ///---------------------------------------------------------
        /// <summary>
        ///     画像表示メイン : 2020/04/14 </summary>
        /// <param name="mImg">
        ///     Mat形式イメージ</param>
        /// <param name="w">
        ///     width</param>
        /// <param name="h">
        ///     height</param>
        ///---------------------------------------------------------
        private void ImgShow(Image mImg, float w, float h)
        {
            int cWidth  = 0;
            int cHeight = 0;

            int pWidth  = panel1.Width  - 2;
            int pHeight = panel1.Height - 2;

            try
            {
                Bitmap bt = new Bitmap(mImg);

                // Bitmapサイズ
                if (pWidth < (bt.Width * w) || pHeight < (bt.Height * h))
                {
                    cWidth  = (int)(bt.Width  * w);
                    cHeight = (int)(bt.Height * h);
                }
                else
                {
                    cWidth  = pWidth;
                    cHeight = pHeight;
                }

                // Bitmap を生成
                Bitmap canvas = new Bitmap(cWidth, cHeight);

                // ImageオブジェクトのGraphicsオブジェクトを作成する
                Graphics g = Graphics.FromImage(canvas);

                // 画像をcanvasの座標(0, 0)の位置に指定のサイズで描画する
                g.DrawImage(bt, 0, 0, bt.Width * w, bt.Height * h);

                //メモリクリア
                bt.Dispose();
                g.Dispose();

                // PictureBox1に表示する
                pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;
                pictureBox1.Image    = canvas;
            }
            catch (Exception ex)
            {
                pictureBox1.Image = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void dg1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < global.flgOff)
            {
                return;
            }

            if (Utility.nulltoStr2(dg1[colTourokuID, e.RowIndex].Value) == global.FLGOFF &&
                Utility.nulltoStr2(dg1[colScanID, e.RowIndex].Value) == global.FLGOFF)
            {
                return;
            }

            string img = "";    // 画像ファイル名

            // 防犯登録IDあり
            if (Utility.nulltoStr2(dg1[colTourokuID, e.RowIndex].Value) != global.FLGOFF)
            {
                // 防犯登録データで検索
                int sId = Utility.StrtoInt(Utility.nulltoStr2(dg1[colTourokuID, e.RowIndex].Value));
                tAdp.FillByID(dts.防犯登録データ, sId);

                foreach (var t in dts.防犯登録データ)
                {
                    img = t.画像名;
                }

                ShowImage(Properties.Settings.Default.imgPath, img);

            }
            else if (Utility.nulltoStr2(dg1[colScanID, e.RowIndex].Value) != global.FLGOFF)
            {
                // SCANIDありSCAN_DATAで検索
                int sId = Utility.StrtoInt(Utility.nulltoStr2(dg1[colScanID, e.RowIndex].Value));
                sAdp.FillByID(dts.SCAN_DATA, sId);

                img = "";

                foreach (var t in dts.SCAN_DATA)
                {
                    img = t.画像名;
                }

                if (img != "")
                {
                    ShowImage(Properties.Settings.Default.scanDataPath, img);
                }
                else
                {
                    // SCAN_DATAに存在しないため防犯登録データで再検索
                    string CPA_Number = global.DATA_CPA + Utility.nulltoStr2(dg1[colNum, e.RowIndex].Value).PadLeft(7, '0');
                    tAdp.FillByNumber(dts.防犯登録データ, CPA_Number);

                    foreach (var t in dts.防犯登録データ)
                    {
                        img = t.画像名;
                    }

                    ShowImage(Properties.Settings.Default.imgPath, img);
                }
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中の登録カード明細をExcel出力します。よろしいですか？",  "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // Excel出力
            Grid2Excel(Properties.Settings.Default.xlsKaishuTemp);
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     回収防犯登録カード明細をExcelに出力 </summary>
        /// <param name="sPath">
        ///     テンプレートシートパス</param>
        ///----------------------------------------------------------------
        private void Grid2Excel(string sPath)
        {
            Cursor = Cursors.WaitCursor;

            try
            {
                IXLWorkbook wb;
                //const string ExcelFilePath = ".\\sample.xlsx";

                // Excelファイルを作る
                using (wb = new XLWorkbook(sPath, XLEventTracking.Disabled))
                {
                    // ワークシートを追加する
                    var ws = wb.Worksheet(1);

                    // 得意先情報をセット
                    ws.Cell(2, 1).SetValue("得意先");
                    ws.Cell(2, 2).SetValue(lblUCode.Text);
                    ws.Cell(2, 3).SetValue(lblUName.Text);

                    int row = 0;

                    // 明細見出し
                    ws.Cell(4, 2).SetValue("登録番号");
                    ws.Cell(4, 3).SetValue("回収");
                    ws.Cell(4, 4).SetValue("回収日");

                    // 登録カード明細情報
                    for (int i = 0; i < dg1.Rows.Count; i++)
                    {
                        row = i + 5;
                        ws.Cell(row, 1).SetValue(Utility.nulltoStr2(dg1[colSeq, i].Value));
                        ws.Cell(row, 2).SetValue(Utility.nulltoStr2(dg1[colNum, i].Value));
                        ws.Cell(row, 3).SetValue(Utility.nulltoStr2(dg1[colStatus, i].Value));
                        ws.Cell(row, 4).SetValue(Utility.nulltoStr2(dg1[colDate, i].Value));

                        ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(row, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(row, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        if (Utility.nulltoStr2(dg1[colStatus, i].Value) == "")
                        { 
                            // 塗りつぶし
                            var sumRangeStyle = ws.Range(row, 1, row, 4).Style;
                            sumRangeStyle.Fill.BackgroundColor = XLColor.MistyRose;
                        }
                    }

                    // 罫線
                    ws.Range(4, 1, row, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Range(4, 1, row, 4).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    // 出庫情報
                    ws.Cell(4, 6).SetValue("出庫日");
                    ws.Cell(5, 6).SetValue("開始番号");
                    ws.Cell(6, 6).SetValue("終了番号");
                    ws.Cell(7, 6).SetValue("出庫部数");
                    ws.Cell(8, 6).SetValue("回収");
                    ws.Cell(9, 6).SetValue("残数");

                    ws.Cell(4, 7).SetValue(lblShukkoDate.Text);
                    ws.Cell(5, 7).SetValue(lblStartNum.Text);
                    ws.Cell(6, 7).SetValue(lblEndNum.Text);
                    ws.Cell(7, 7).SetValue(lblSuu.Text);
                    ws.Cell(8, 7).SetValue(lblKaishu.Text);
                    ws.Cell(9, 7).SetValue(lblZan.Text);

                    // 罫線
                    ws.Range(4, 6, 9, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Range(4, 6, 9, 7).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    // セルに書式設定・塗りつぶし
                    var sumCellStyle = ws.Cell("A2").Style;
                    sumCellStyle.Fill.BackgroundColor = XLColor.LightGray;

                    sumCellStyle = ws.Range(4,1,4,4).Style;
                    sumCellStyle.Fill.BackgroundColor = XLColor.LightGray;

                    sumCellStyle = ws.Range(4, 6, 4, 7).Style;
                    sumCellStyle.Fill.BackgroundColor = XLColor.LightGray;

                    sumCellStyle = ws.Range(5, 6, 9, 6).Style;
                    sumCellStyle.Fill.BackgroundColor = XLColor.LightGray;

                    //ws.Cell(7, 7).FormulaA1 = "SUM(A1:A2)";

                    //// セルに書式設定
                    //var sumCellStyle = ws.Cell("A3").Style;
                    //sumCellStyle.Fill.BackgroundColor = XLColor.Red; // 塗りつぶし
                    //sumCellStyle.NumberFormat.Format = "#,##0.00"; // 数値の書式
                    //                                               // 次のようにメソッドチェーンでも書ける
                    //                                               //worksheet.Cell("A3").SetFormulaA1("SUM(A1:A2)")
                    //                                               //                    .Style.Fill.SetBackgroundColor(XLColor.Red)
                    //                                               //                          .NumberFormat.SetFormat("#,##0.00");





                    //ダイアログボックスの初期設定
                    DialogResult ret;
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog
                    {
                        Title = "回収防犯登録カード明細",
                        OverwritePrompt = true,
                        RestoreDirectory = true,
                        FileName = lblUName.Text + " 防犯登録カード回収状況",
                        Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*"
                    };

                    //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileDialog1.FileName;

                        // ワークブックを保存する
                        wb.SaveAs(fileName);
                    }
                }

                MessageBox.Show("Excel出力終了しました");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
    }
}
