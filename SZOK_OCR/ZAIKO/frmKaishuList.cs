using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Windows.Forms;
using SZOK_OCR.Common;

namespace SZOK_OCR.ZAIKO
{
    public partial class frmKaishuList : Form
    {
        public frmKaishuList()
        {
            InitializeComponent();
        }

        cardDataSet dts = new cardDataSet();
        cardDataSetTableAdapters.回収データTableAdapter kAdp = new cardDataSetTableAdapters.回収データTableAdapter();
        cardDataSetTableAdapters.出庫データTableAdapter sAdp = new cardDataSetTableAdapters.出庫データTableAdapter();

        private void frmKaishuList_Load(object sender, EventArgs e)
        {
            GridViewSetting(dg1);

            btnExcel.Enabled = false;
        }

        // データグリッドビューカラム定義
        string colUCode    = "col1";
        string colUName    = "col2";
        string colDate     = "col3";
        string colMikaishu = "col4";
        string colSNum     = "col5";
        string colENum     = "col6";
        string colID       = "col7";  // 2021/10/22


        ///----------------------------------------------------------
        /// <summary>
        ///     未回収防犯登録カードリスト表示 </summary>
        ///----------------------------------------------------------
        private void DataShow()
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                DateTime dt_s = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day, 0, 0, 0);
                DateTime dt_e = new DateTime(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day, 23, 59, 59);

                sAdp.FillByShukkoDayRange(dts.出庫データ, dt_s, dt_e);

                IEnumerable<cardDataSet.出庫データRow> shukko = null;

                if (txtUser.Text.Trim() != string.Empty)
                {
                    // 指定得意先名
                    shukko = dts.出庫データ.Where(a => a.店名.Contains(txtUser.Text.Trim())).OrderBy(a => a.出庫日).ThenBy(a => a.店番);
                }
                else
                {
                    shukko = dts.出庫データ.OrderBy(a => a.出庫日).ThenBy(a => a.店番);
                }

                // グリッドビュー初期化
                dg1.Rows.Clear();

                foreach (var t in shukko)
                {
                    kAdp.FillByKaishu(dts.回収データ, t.ID, dt_e);

                    for (int i = t.開始登録番号; i <= t.終了登録番号; i++)
                    {
                        if (dts.回収データ.Any(a => a.登録番号 == i))
                        {
                            // 回収済みはネグる
                            continue;
                        }

                        dg1.Rows.Add();
                        dg1[colDate, dg1.RowCount - 1].Value     = t.出庫日.ToShortDateString();
                        dg1[colUCode, dg1.RowCount - 1].Value    = t.店番;
                        dg1[colUName, dg1.RowCount - 1].Value    = t.店名;
                        dg1[colSNum, dg1.RowCount - 1].Value     = t.開始登録番号;
                        dg1[colENum, dg1.RowCount - 1].Value     = t.終了登録番号;
                        dg1[colMikaishu, dg1.RowCount - 1].Value = i;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                Cursor = Cursors.Default;

                if (dg1.RowCount > 0)
                {
                    btnExcel.Enabled = true;
                    dg1.CurrentCell  = null;
                }
                else
                {
                    btnExcel.Enabled = false;
                }
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///--------------------------------------------------------------------
        private void GridViewSetting(DataGridView tempDGV)
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
                tempDGV.RowTemplate.Height  = 20;

                // 全体の高さ
                tempDGV.Height = 482;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                //各列幅指定
                tempDGV.Columns.Add(colDate,     "出庫日");
                tempDGV.Columns.Add(colUCode,    "コード");
                tempDGV.Columns.Add(colUName,    "得意先名");
                tempDGV.Columns.Add(colSNum,     "開始番号");
                tempDGV.Columns.Add(colENum,     "終了番号");
                tempDGV.Columns.Add(colMikaishu, "未回収番号");

                tempDGV.Columns[colDate].Width     = 110;
                tempDGV.Columns[colUCode].Width    = 90;
                tempDGV.Columns[colSNum].Width     = 110;
                tempDGV.Columns[colENum].Width     = 110;
                tempDGV.Columns[colMikaishu].Width = 110;

                tempDGV.Columns[colUName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[colDate].DefaultCellStyle.Alignment     = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colUCode].DefaultCellStyle.Alignment    = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colMikaishu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colSNum].DefaultCellStyle.Alignment     = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colENum].DefaultCellStyle.Alignment     = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colMikaishu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = true;

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

        private void button1_Click(object sender, EventArgs e)
        {
            DataShow();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmKaishuList_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中の未回収カードリストをExcel出力します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            Grid2Excel(Properties.Settings.Default.xlsMikaishuList);
        }

        private void Grid2Excel(string sPath)
        {
            Cursor = Cursors.WaitCursor;

            try
            {
                IXLWorkbook wb;

                // Excelファイルを読み込む
                using (wb = new XLWorkbook(sPath, XLEventTracking.Disabled))
                {
                    // ワークシートを追加する
                    var ws = wb.Worksheet(1);

                    // 指定得意先・集計期間をセット
                    ws.Cell(2, 2).SetValue(txtUser.Text.Trim());
                    ws.Cell(2, 4).SetValue("集計期間");
                    ws.Cell(2, 5).SetValue(dateTimePicker1.Value.ToShortDateString() + "～" + dateTimePicker2.Value.ToShortDateString());

                    int row = 0;

                    // 登録カード明細情報
                    for (int i = 0; i < dg1.Rows.Count; i++)
                    {
                        row = i + 5;
                        ws.Cell(row, 1).SetValue(Utility.nulltoStr2(dg1[colDate,  i].Value));
                        ws.Cell(row, 2).SetValue(Utility.nulltoStr2(dg1[colUCode, i].Value));
                        ws.Cell(row, 3).SetValue(Utility.nulltoStr2(dg1[colUName, i].Value));
                        ws.Cell(row, 4).SetValue(Utility.nulltoStr2(dg1[colSNum,  i].Value));
                        ws.Cell(row, 5).SetValue(Utility.nulltoStr2(dg1[colENum,  i].Value));
                        ws.Cell(row, 6).SetValue(Utility.nulltoStr2(dg1[colMikaishu, i].Value));

                        ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(row, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(row, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(row, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    }

                    // 罫線
                    ws.Range(4, 1, row, 6).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    ws.Range(4, 1, row, 6).Style.Border.InsideBorder  = XLBorderStyleValues.Thin;

                    // セルに書式設定・塗りつぶし
                    var sumCellStyle = ws.Cell("A2").Style;
                    sumCellStyle.Fill.BackgroundColor = XLColor.LightGray;

                    sumCellStyle = ws.Cell("D2").Style;
                    sumCellStyle.Fill.BackgroundColor = XLColor.LightGray;

                    sumCellStyle = ws.Range(4, 1, 4, 6).Style;
                    sumCellStyle.Fill.BackgroundColor = XLColor.LightGray;


                    //ダイアログボックスの初期設定
                    DialogResult ret;
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog
                    {
                        Title = "未回収カードリスト",
                        OverwritePrompt  = true,
                        RestoreDirectory = true,
                        FileName = "未回収カードリスト",
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
