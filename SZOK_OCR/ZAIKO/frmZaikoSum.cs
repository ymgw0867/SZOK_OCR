﻿using System;
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
    public partial class frmZaikoSum : Form
    {
        public frmZaikoSum()
        {
            InitializeComponent();

            // データ読み込み
            sAdp.Fill(dts.出庫データ);
            kAdp.Fill(dts.回収データ);
        }

        cardDataSet dts = new cardDataSet();
        cardDataSetTableAdapters.出庫データTableAdapter sAdp = new cardDataSetTableAdapters.出庫データTableAdapter();
        cardDataSetTableAdapters.回収データTableAdapter kAdp = new cardDataSetTableAdapters.回収データTableAdapter();

        private void frmZaikoSum_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Checked = false;
            txtUser.Text = string.Empty;

            // データグリッドビュー定義
            gridViewSetting(dataGridView1);

            comboBox1.SelectedIndex = 0;
            button2.Enabled = false;
        }

        // データグリッドビューカラム定義
        string colUCode = "col1";
        string colUName = "col2";
        string colDate = "col3";
        string colShukko = "col4";
        string colKaishu = "col5";
        string colZansu = "col6";
        string colSNum = "col7";
        string colENum = "col8";

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
                tempDGV.Height = 502;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                //各列幅指定
                tempDGV.Columns.Add(colUCode, "コード");
                tempDGV.Columns.Add(colUName, "得意先名");
                tempDGV.Columns.Add(colDate, "出庫日");
                tempDGV.Columns.Add(colSNum, "開始番号");
                tempDGV.Columns.Add(colENum, "終了番号");
                tempDGV.Columns.Add(colShukko, "出庫部数");
                tempDGV.Columns.Add(colKaishu, "回収");
                tempDGV.Columns.Add(colZansu, "残数");

                tempDGV.Columns[colUCode].Width = 90;
                tempDGV.Columns[colUName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[colDate].Width = 110;
                tempDGV.Columns[colSNum].Width = 110;
                tempDGV.Columns[colENum].Width = 110;
                tempDGV.Columns[colShukko].Width = 80;
                tempDGV.Columns[colKaishu].Width = 80;
                tempDGV.Columns[colZansu].Width = 80;

                tempDGV.Columns[colUCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colShukko].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colSNum].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colENum].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colKaishu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[colZansu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

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

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                ZaikoSummary(dataGridView1);
            }
            else
            {
                ZaikoSummaryTotal(dataGridView1);
            }
        }

        private void ZaikoSummary(DataGridView g)
        {
            Cursor = Cursors.WaitCursor;

            int Tenban = 0;
            string tenName = string.Empty;

            int[] ShukkoTl = new int[2];
            int[] KaishuTl = new int[2];
            for (int i = 0; i < 2; i++)
            {
                ShukkoTl[i] = 0;
                KaishuTl[i] = 0;
            }

            try
            {
                int iX = 0;

                var s = dts.出庫データ.OrderBy(a => a.店番).ThenBy(a => a.出庫日);

                // 出庫基準年月日
                if (dateTimePicker1.Checked)
                {
                    DateTime dt = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
                    s = s.Where(a => a.出庫日 >= dt).OrderBy(a => a.店番).ThenBy(a => a.出庫日);
                }

                // 指定得意先名
                if (txtUser.Text.Trim() != string.Empty)
                {
                    s = s.Where(a => a.店名.Contains(txtUser.Text)).OrderBy(a => a.店番).ThenBy(a => a.出庫日);
                }

                dataGridView1.Rows.Clear();

                foreach (var t in s)
                {
                    if (Tenban != 0)
                    {
                        if (Tenban != t.店番)
                        {
                            dataGridView1.Rows.Add();

                            g[colUCode, iX].Value = "";
                            g[colUName, iX].Value = tenName + "　合計";
                            g[colDate, iX].Value = "";
                            g[colSNum, iX].Value = "";
                            g[colENum, iX].Value = "";
                            g[colShukko, iX].Value = ShukkoTl[0].ToString("#,##0");
                            g[colKaishu, iX].Value = KaishuTl[0].ToString("#,##0");
                            g[colZansu, iX].Value = (ShukkoTl[0] - KaishuTl[0]).ToString("#,##0");

                            ShukkoTl[0] = 0;
                            KaishuTl[0] = 0;
                            iX++;
                        }
                    }

                    dataGridView1.Rows.Add();

                    g[colUCode, iX].Value = t.店番;
                    g[colUName, iX].Value = t.店名;
                    g[colDate, iX].Value = t.出庫日.ToShortDateString();
                    g[colSNum, iX].Value = t.開始登録番号;
                    g[colENum, iX].Value = t.終了登録番号;
                    g[colShukko, iX].Value = t.部数.ToString("#,##0");
                    int kaishu = t.Get回収データRows().Count();
                    g[colKaishu, iX].Value = kaishu.ToString("#,##0");
                    g[colZansu, iX].Value = (t.部数 - kaishu).ToString("#,##0");

                    for (int i = 0; i < 2; i++)
                    {
                        ShukkoTl[i] += t.部数;
                        KaishuTl[i] += kaishu;
                    }

                    Tenban = t.店番;
                    tenName = t.店名;

                    iX++;
                }

                for (int i = 0; i < 2; i++)
                {
                    dataGridView1.Rows.Add();

                    g[colUCode, iX].Value = "";

                    if (i == 0)
                    {
                        g[colUName, iX].Value = tenName + "　合計";
                    }
                    else
                    {
                        g[colUName, iX].Value = "総計";
                    }

                    g[colDate, iX].Value = "";
                    g[colSNum, iX].Value = "";
                    g[colENum, iX].Value = "";
                    g[colShukko, iX].Value = ShukkoTl[i].ToString("#,##0");
                    g[colKaishu, iX].Value = KaishuTl[i].ToString("#,##0");
                    g[colZansu, iX].Value = (ShukkoTl[i] - KaishuTl[i]).ToString("#,##0");

                    iX++;
                }

                if (g.Rows.Count > 0)
                {
                    g.CurrentCell = null;
                    button2.Enabled = true;
                }
                else
                {
                    button2.Enabled = false;
                }
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


        private void ZaikoSummaryTotal(DataGridView g)
        {
            Cursor = Cursors.WaitCursor;

            int Tenban = 0;
            string tenName = string.Empty;

            int[] ShukkoTl = new int[2];
            int[] KaishuTl = new int[2];
            for (int i = 0; i < 2; i++)
            {
                ShukkoTl[i] = 0;
                KaishuTl[i] = 0;
            }

            try
            {
                int iX = 0;

                var s = dts.出庫データ.OrderBy(a => a.店番).ThenBy(a => a.出庫日);

                // 出庫基準年月日
                if (dateTimePicker1.Checked)
                {
                    DateTime dt = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
                    s = s.Where(a => a.出庫日 >= dt).OrderBy(a => a.店番).ThenBy(a => a.出庫日);
                }

                // 指定得意先名
                if (txtUser.Text.Trim() != string.Empty)
                {
                    s = s.Where(a => a.店名.Contains(txtUser.Text)).OrderBy(a => a.店番).ThenBy(a => a.出庫日);
                }

                dataGridView1.Rows.Clear();

                foreach (var t in s)
                {
                    if (Tenban != 0)
                    {
                        if (Tenban != t.店番)
                        {
                            dataGridView1.Rows.Add();

                            g[colUCode, iX].Value = Tenban;
                            g[colUName, iX].Value = tenName;
                            g[colDate, iX].Value = "";
                            g[colSNum, iX].Value = "";
                            g[colENum, iX].Value = "";
                            g[colShukko, iX].Value = ShukkoTl[0].ToString("#,##0");
                            g[colKaishu, iX].Value = KaishuTl[0].ToString("#,##0");
                            g[colZansu, iX].Value = (ShukkoTl[0] - KaishuTl[0]).ToString("#,##0");

                            ShukkoTl[0] = 0;
                            KaishuTl[0] = 0;
                            iX++;
                        }
                    }

                    //dataGridView1.Rows.Add();

                    //g[colUCode, iX].Value = t.店番;
                    //g[colUName, iX].Value = t.店名;
                    //g[colDate, iX].Value = t.出庫日.ToShortDateString();
                    //g[colSNum, iX].Value = t.開始登録番号;
                    //g[colENum, iX].Value = t.終了登録番号;
                    //g[colShukko, iX].Value = t.部数.ToString("#,##0");
                    //int kaishu = t.Get回収データRows().Count();
                    //g[colKaishu, iX].Value = kaishu.ToString("#,##0");
                    //g[colZansu, iX].Value = (t.部数 - kaishu).ToString("#,##0");

                    for (int i = 0; i < 2; i++)
                    {
                        ShukkoTl[i] += t.部数;
                        KaishuTl[i] += t.Get回収データRows().Count();
                    }

                    Tenban = t.店番;
                    tenName = t.店名;

                    //iX++;
                }

                for (int i = 0; i < 2; i++)
                {
                    dataGridView1.Rows.Add();


                    if (i == 0)
                    {
                        g[colUCode, iX].Value = Tenban;
                        g[colUName, iX].Value = tenName;
                    }
                    else
                    {
                        g[colUCode, iX].Value = "";
                        g[colUName, iX].Value = "総計";
                    }

                    g[colDate, iX].Value = "";
                    g[colSNum, iX].Value = "";
                    g[colENum, iX].Value = "";
                    g[colShukko, iX].Value = ShukkoTl[i].ToString("#,##0");
                    g[colKaishu, iX].Value = KaishuTl[i].ToString("#,##0");
                    g[colZansu, iX].Value = (ShukkoTl[i] - KaishuTl[i]).ToString("#,##0");

                    iX++;
                }

                if (g.Rows.Count > 0)
                {
                    g.CurrentCell = null;
                    button2.Enabled = true;
                }
                else
                {
                    button2.Enabled = false;
                }
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


        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmZaikoSum_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelOutput(Properties.Settings.Default.xlsZaiko, dataGridView1);
        }

        private void ExcelOutput(string sPath, DataGridView g)
        {
            Cursor = Cursors.WaitCursor;

            try
            {
                using (var bk = new XLWorkbook(XLEventTracking.Disabled))
                {
                    // ワークシートを作成
                    bk.Style.Font.FontName = "ＭＳ ゴシック";
                    bk.Style.Font.FontSize = 11;

                    var sheet1 = bk.AddWorksheet("在庫集計表");

                    sheet1.Range("A1:C1").Merge();

                    string kijun = "出庫基準日：";

                    if (dateTimePicker1.Checked)
                    {
                        kijun += dateTimePicker1.Value.ToShortDateString() + "～";
                    }
                    else
                    {
                        kijun += "全期間";
                    }

                    sheet1.Cell("A1").SetValue(kijun);

                    sheet1.Cell("A2").SetValue("コード");
                    sheet1.Cell("B2").SetValue("得意先名");
                    sheet1.Cell("C2").SetValue("出庫日");
                    sheet1.Cell("D2").SetValue("開始番号");
                    sheet1.Cell("E2").SetValue("終了番号");
                    sheet1.Cell("F2").SetValue("出庫部数");
                    sheet1.Cell("G2").SetValue("回収");
                    sheet1.Cell("H2").SetValue("残数");

                    for (int i = 0; i < g.Rows.Count; i++)
                    {
                        sheet1.Cell(i + 3, 1).Value = g[colUCode, i].Value.ToString();
                        sheet1.Cell(i + 3, 2).Value = g[colUName, i].Value.ToString();
                        sheet1.Cell(i + 3, 3).Value = g[colDate, i].Value.ToString();
                        sheet1.Cell(i + 3, 4).Value = g[colSNum, i].Value.ToString();
                        sheet1.Cell(i + 3, 5).Value = g[colENum, i].Value.ToString();
                        sheet1.Cell(i + 3, 6).Value = g[colShukko, i].Value.ToString().Replace(",", "");
                        sheet1.Cell(i + 3, 7).Value = g[colKaishu, i].Value.ToString().Replace(",", "");
                        sheet1.Cell(i + 3, 8).Value = g[colZansu, i].Value.ToString().Replace(",", "");

                        if (g[colUName, i].Value.ToString().Contains("合計"))
                        {
                            sheet1.Range("A" + (i + 3) + ":H" + (i + 3)).Style.Fill.BackgroundColor = XLColor.LightGray;
                            sheet1.Range("A" + (i + 3) + ":H" + (i + 3)).Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
                            sheet1.Range("A" + (i + 3) + ":H" + (i + 3)).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
                        }
                    }

                    sheet1.Column("A").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet1.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    sheet1.Column("C").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet1.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet1.Column("E").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet1.Column("F").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    sheet1.Column("G").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    sheet1.Column("H").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    sheet1.Column("C").Style.NumberFormat.SetFormat("yyyy/mm/dd");
                    sheet1.Column("F").Style.NumberFormat.SetFormat("#,##0");
                    sheet1.Column("G").Style.NumberFormat.SetFormat("#,##0");
                    sheet1.Column("H").Style.NumberFormat.SetFormat("#,##0");

                    sheet1.Column("A").Width = 10;
                    sheet1.Column("B").Width = 50;
                    sheet1.Column("C").Width = 16;
                    sheet1.Column("D").Width = 12;
                    sheet1.Column("E").Width = 12;
                    sheet1.Column("F").Width = 12;
                    sheet1.Column("G").Width = 10;
                    sheet1.Column("H").Width = 10;

                    sheet1.Column("B").Style.Alignment.ShrinkToFit = true;

                    sheet1.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    sheet1.Row(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet1.Range("A2:H2").Style.Fill.BackgroundColor = XLColor.LightGray;

                    // 全体の罫線
                    sheet1.Range("A2:H2").Style.Border.SetTopBorder(XLBorderStyleValues.Thin);
                    sheet1.Range("A2:H2").Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
                    sheet1.Range(sheet1.Cell("A2"), sheet1.LastCellUsed()).Style.Border.SetLeftBorder(XLBorderStyleValues.Thin)
                        .Border.SetRightBorder(XLBorderStyleValues.Thin);
                    sheet1.Range("A" + (sheet1.RowsUsed().Count()) + ":H" + (sheet1.RowsUsed().Count())).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);
                    
                    DialogResult ret;

                    string fName = "";

                    if (comboBox1.SelectedIndex == 0)
                    {
                        fName = "得意先出庫別在庫集計表";
                    }
                    else
                    {
                        fName = "得意先別在庫集計表";
                    }

                    //ダイアログボックスの初期設定
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Title = "在庫集計表";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = DateTime.Today.Year + DateTime.Today.Month.ToString("D2") + DateTime.Today.Day.ToString("D2") + " " + fName;
                    saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                    //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        // エクセル保存
                        fileName = saveFileDialog1.FileName;
                        bk.SaveAs(fileName);

                        // メッセージ
                        MessageBox.Show("Excel出力が終了しました", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
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
    }
}