using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SZOK_OCR.Common;

namespace SZOK_OCR.OCR
{
    public partial class frmZipCode : Form
    {
        public frmZipCode(string zipCode)
        {
            InitializeComponent();
            _zipCode = zipCode;
        }

        string _zipCode = string.Empty;
        string[] zipArray = null;

        string colZipCode = "col1";
        string colZipAdd1 = "col2";
        string colZipAdd2 = "col3";

        private void frmZipCode_Load(object sender, EventArgs e)
        {
            // 郵便番号CSVデータを配列に読み込み
            Utility.zipCsvLoad(ref zipArray);

            // データグリッドビュー定義
            gridViewSetting(dg1);
            gridViewSetting(dg2);
            
            txtZipCode.Text = _zipCode;
            txtZipCode.ImeMode = ImeMode.Off;
            txtAddress.Text = string.Empty;
            txtAddress.ImeMode = ImeMode.KatakanaHalf;

            button1.Enabled = false;
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
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 11, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 11, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 322;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                //各列幅指定
                tempDGV.Columns.Add(colZipCode, "〒");
                tempDGV.Columns.Add(colZipAdd1, "住所");
                tempDGV.Columns.Add(colZipAdd2, "カナ");
                
                tempDGV.Columns[colZipCode].Width = 100;
                tempDGV.Columns[colZipAdd1].Width = 300;
                tempDGV.Columns[colZipAdd2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[colZipCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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

        ///----------------------------------------------------------------
        /// <summary>
        ///     郵便番号から住所を取得する </summary>
        /// <param name="sAdd">
        ///     郵便番号</param>
        /// <param name="z">
        ///     郵便番号配列</param>
        ///----------------------------------------------------------------
        private void getZipToAdd(string zCode, string [] z)
        {
            if (z != null)
            {
                int zLen = zCode.Length;    // 検索文字列文字数取得
                int iX = 0;
                dg1.Rows.Clear();

                if (zLen == 0)
                {
                    return;
                }

                this.Cursor = Cursors.WaitCursor;

                foreach (var t in z)
                {
                    string[] zip = t.Split(',');
                    string zipMs = zip[2].Replace("\"", "");

                    if (zipMs.Length < zLen)
                    {
                        continue;
                    }

                    if (zipMs.Substring(0, zLen) == zCode)
                    {
                        dg1.Rows.Add();

                        dg1[colZipCode, iX].Value = zip[2].Replace("\"", "");
                        dg1[colZipAdd1, iX].Value = (zip[7] + zip[8]).Replace("\"", "");
                        dg1[colZipAdd2, iX].Value = (zip[4] + " " + zip[5]).Replace("\"", "");

                        iX++;
                    }
                }

                this.Cursor = Cursors.Default;

                if (iX > 0)
                {
                    dg1.CurrentCell = null;
                    dg1.Sort(dg1.Columns[colZipCode], ListSortDirection.Ascending);
                }
            }
        }

        private void txtZipCode_TextChanged(object sender, EventArgs e)
        {
            getZipToAdd(txtZipCode.Text, zipArray);
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     住所から郵便番号を取得する </summary>
        /// <param name="sAdd">
        ///     住所</param>
        /// <param name="z">
        ///     郵便番号配列</param>
        ///----------------------------------------------------------------
        private void getAddToZip(string sAdd, string[] z)
        {
            System.Globalization.CompareInfo ci = System.Globalization.CultureInfo.CurrentCulture.CompareInfo;

            if (z != null)
            {
                int iX = 0;
                dg2.Rows.Clear();
                //bool cmp = true;
                
                this.Cursor = Cursors.WaitCursor;

                foreach (var t in z)
                {
                    string[] zip = t.Split(',');
                    string zipMs = (zip[4] + zip[5]).Replace("\"", "") + " " +
                                   (zip[7] + zip[8]).Replace("\"", "");

                    //// 大文字小文字の違いを無視して比較する
                    //if (zipMs.IndexOf(sAdd, StringComparison.OrdinalIgnoreCase) == -1)
                    //{
                    //    //ひらがなとカタカナを区別しない
                    //    if (ci.IndexOf(zipMs, sAdd, System.Globalization.CompareOptions.IgnoreKanaType) == -1)
                    //    {
                    //        //半角と全角を区別しない
                    //        if (ci.IndexOf(zipMs, sAdd, System.Globalization.CompareOptions.IgnoreWidth) == -1)
                    //        {
                    //            cmp = false;
                    //        }
                    //    }
                    //}


                    sAdd = Utility.getStrConv(sAdd);

                    if (zipMs.Contains(sAdd))
                    {
                        dg2.Rows.Add();

                        dg2[colZipCode, iX].Value = zip[2].Replace("\"", "");
                        dg2[colZipAdd1, iX].Value = (zip[7] + zip[8]).Replace("\"", "");
                        dg2[colZipAdd2, iX].Value = (zip[4] + " " + zip[5]).Replace("\"", "");

                        iX++;
                    }
                }

                this.Cursor = Cursors.Default;

                if (iX > 0)
                {
                    dg2.CurrentCell = null;
                    dg2.Sort(dg2.Columns[colZipCode], ListSortDirection.Ascending);
                }
            }
        }

        private void txtAddress_TextChanged(object sender, EventArgs e)
        {
            getAddToZip(txtAddress.Text, zipArray);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                zipSelect(dg1);
            } 
            else if (tabControl1.SelectedIndex == 1)
            {
                zipSelect(dg2);
            }

        }

        // データ作成画面に渡す情報
        public string rZipCode { get; set; }    // 郵便番号
        public string rAddFuri { get; set; }    // 住所フリガナ
        public string rAdd { get; set; }        // 住所

        ///--------------------------------------------------------------
        /// <summary>
        ///     郵便番号・住所選択 </summary>
        /// <param name="d">
        ///     データグリッドビューオブジェクト</param>
        ///--------------------------------------------------------------
        private void zipSelect(DataGridView d)
        {
            if (d.SelectedRows.Count == 0)
            {
                return;
            }

            rZipCode = d[colZipCode, d.SelectedRows[0].Index].Value.ToString();
            rAdd = d[colZipAdd1, d.SelectedRows[0].Index].Value.ToString();
            rAddFuri = d[colZipAdd2, d.SelectedRows[0].Index].Value.ToString();

            // 閉じる
            this.Close();
        }

        private void frmZipCode_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void dg1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            button1.Enabled = true;
        }

        private void dg2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            button1.Enabled = true;
        }

        private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            rZipCode = string.Empty;
            rAdd = string.Empty;
            rAddFuri = string.Empty;

            // 閉じる
            this.Close();
        }
    }
}
