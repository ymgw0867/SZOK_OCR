using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SZOK_OCR.Common;

namespace SZOK_OCR.DATA
{
    public partial class frmScanList : Form
    {
        public frmScanList()
        {
            InitializeComponent();
        }

        cardDataSet dts = new cardDataSet();
        cardDataSetTableAdapters.SCAN_DATATableAdapter adp = new cardDataSetTableAdapters.SCAN_DATATableAdapter();

        global g = new global();

        // データグリッドビューカラム定義
        string coldKbn = "col1";
        string colCPA = "col2";
        string colCarbodyNum = "col3";
        string colyymmdd = "col4";
        string colMaker = "col5";
        string colColor = "col6";
        string colCarStyle = "col7";
        string colSharyoNum = "col8";
        string colCarName = "col9";
        string colZip = "col10";
        string colAdd = "col11";
        string colFuri = "col12";
        string colTel = "col13";
        string colID = "colID";
        string colCsv = "col14";
        string colJyogai = "col15";
        string colLabel = "col16";
        string colName = "col17";

        private void frmCardList_Load(object sender, EventArgs e)
        {
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッドビュー定義
            gridViewSetting(dg);

            // 車種コンボアイテムセット
            getCarStyleArray();

            // 検索欄初期化
            dispInitial();
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     車種コンボボックスに値をセットする </summary>
        ///-----------------------------------------------------------------
        private void getCarStyleArray()
        {
            cmbCarStyle.Items.Add("全て");

            global g = new global();
            for (int i = 0; i < g.arrStyle.GetLength(0); i++)
            {
                cmbCarStyle.Items.Add(g.arrStyle[i, 0] + ":" + g.arrStyle[i, 1]);
            }
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        ///-----------------------------------------------------------------
        private void dispInitial()
        {
            cmbShubetsu.SelectedIndex = 0;
            txtsCpa.Text = string.Empty;
            txtsCarbodyNum.Text = string.Empty;
            txtsYY.Text = string.Empty;
            txtsMM.Text = string.Empty;
            txtsDD.Text = string.Empty;
            txtsMaker.Text = string.Empty;
            txtsColor.Text = string.Empty;
            cmbCarStyle.SelectedIndex = 0;
            txtsSharyoNum.Text = string.Empty;
            txtsCarName.Text = string.Empty;
            txtsZip1.Text = string.Empty;
            txtsZip2.Text = string.Empty;
            txtsAdd.Text = string.Empty;
            txtsFuri.Text = string.Empty;
            txtsTel1.Text = string.Empty;
            txtsTel2.Text = string.Empty;
            txtsTel3.Text = string.Empty;
            linkLabel2.Enabled = false;
            txtLabel.Text = string.Empty;
            txtName.Text = string.Empty;
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
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 10, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 10, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 542;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                //各列幅指定
                tempDGV.Columns.Add(coldKbn, "種別");
                tempDGV.Columns.Add(colCPA, "登録番号");
                tempDGV.Columns.Add(colCarbodyNum, "車体番号");
                tempDGV.Columns.Add(colyymmdd, "登録年月日");
                tempDGV.Columns.Add(colMaker, "メーカー");
                tempDGV.Columns.Add(colColor, "カラー");
                tempDGV.Columns.Add(colCarStyle, "車種");
                tempDGV.Columns.Add(colSharyoNum, "車両番号");
                tempDGV.Columns.Add(colCarName, "車名");
                tempDGV.Columns.Add(colZip, "〒");
                tempDGV.Columns.Add(colAdd, "住所");
                tempDGV.Columns.Add(colFuri, "氏名");
                tempDGV.Columns.Add(colTel, "ＴＥＬ／携帯");
                //tempDGV.Columns.Add(colCsv, "静岡県警用CSV作成");
                //tempDGV.Columns.Add(colJyogai, "除外");
                tempDGV.Columns.Add(colLabel, "ラベル名");
                tempDGV.Columns.Add(colName, "処理担当者");
                tempDGV.Columns.Add(colID, "");

                tempDGV.Columns[colID].Visible = false;

                tempDGV.Columns[coldKbn].Width = 90;
                tempDGV.Columns[colCPA].Width = 140;
                tempDGV.Columns[colCarbodyNum].Width = 260;
                tempDGV.Columns[colyymmdd].Width = 110;
                tempDGV.Columns[colMaker].Width = 100;
                tempDGV.Columns[colColor].Width = 100;
                tempDGV.Columns[colCarStyle].Width = 100;
                tempDGV.Columns[colSharyoNum].Width = 100;
                tempDGV.Columns[colCarName].Width = 100;
                tempDGV.Columns[colZip].Width = 100;
                tempDGV.Columns[colAdd].Width = 300;
                tempDGV.Columns[colFuri].Width = 120;
                tempDGV.Columns[colTel].Width = 120;

                //tempDGV.Columns[colCsv].Width = 170;
                //tempDGV.Columns[colJyogai].Width = 60;

                tempDGV.Columns[colLabel].Width = 140;
                tempDGV.Columns[colName].Width = 100;

                tempDGV.Columns[colyymmdd].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colZip].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //tempDGV.Columns[colCsv].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //tempDGV.Columns[colJyogai].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colLabel].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

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
            if (Utility.StrtoInt(txtsYY.Text) == global.flgOff)
            {
                MessageBox.Show("登録年を必ず指定してください", "検索項目指定", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // リスト表示
            dataShow();
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     データグリッドに防犯登録カードデータを表示する </summary>
        /// <returns>
        ///     表示件数</returns>
        ///-------------------------------------------------------------------------
        private int dataShow()
        {
            label22.Text = string.Empty;    // 2019/11/15

            // 2019/06/25
            dg.Rows.Clear();

            // 2019/06/25
            System.Threading.Thread.Sleep(1000);
            Application.DoEvents();

            // 2019/06/25
            this.Cursor = Cursors.WaitCursor;

            // 2019/11/15
            adp.FillByYear(dts.SCAN_DATA, txtsYY.Text);


            var s = dts.SCAN_DATA.OrderBy(q => q.登録番号);

            // データ種別
            if (cmbShubetsu.SelectedIndex > 0)
            {
                s = s.Where(q => q.データ区分 == cmbShubetsu.SelectedIndex - 1).OrderBy(q => q.登録番号);
            }

            // 登録番号
            if (txtsCpa.Text != string.Empty)
            {
                s = s.Where(q => q.登録番号.Contains(txtsCpa.Text)).OrderBy(q => q.登録番号);
            }

            // 車体番号
            if (txtsCarbodyNum.Text != string.Empty)
            {
                s = s.Where(q => q.車体番号.Contains(txtsCarbodyNum.Text)).OrderBy(q => q.登録番号);
            }

            // 登録年
            if (txtsYY.Text != string.Empty)
            {
                s = s.Where(q => q.登録年.ToString() == txtsYY.Text).OrderBy(q => q.登録番号);
            }

            // 登録月
            if (txtsMM.Text != string.Empty)
            {
                s = s.Where(q => q.登録月.ToString() == txtsMM.Text).OrderBy(q => q.登録番号);
            }

            // 登録日
            if (txtsDD.Text != string.Empty)
            {
                s = s.Where(q => q.登録日.ToString() == txtsDD.Text).OrderBy(q => q.登録番号);
            }

            // メーカー
            if (txtsMaker.Text != string.Empty)
            {
                s = s.Where(q => q.メーカー.Contains(Utility.getStrConv(txtsMaker.Text))).OrderBy(q => q.登録番号);
            }

            // カラー
            if (txtsColor.Text != string.Empty)
            {
                s = s.Where(q => q.塗色.Contains(Utility.getStrConv(txtsColor.Text))).OrderBy(q => q.登録番号);
            }

            // 車種
            if (cmbCarStyle.SelectedIndex > 0)
            {
                string cs = cmbCarStyle.Text.Substring(0, 2);

                s = s.Where(q => q.車種.ToString().PadLeft(2, '0') == cs).OrderBy(q => q.登録番号);
            }

            // 車両番号
            if (txtsSharyoNum.Text != string.Empty)
            {
                s = s.Where(q => !q.Is車両番号1Null() && (q.車両番号1 + q.車両番号2).Contains(txtsSharyoNum.Text)).OrderBy(q => q.登録番号);
            }

            // 車名
            if (txtsCarName.Text != string.Empty)
            {
                s = s.Where(q => q.車名.Contains(Utility.getStrConv(txtsCarName.Text))).OrderBy(q => q.登録番号);
            }

            // 郵便番号
            if (txtsZip1.Text != string.Empty)
            {
                s = s.Where(q => q.郵便番号1.Contains(txtsZip1.Text)).OrderBy(q => q.登録番号);
            }

            if (txtsZip2.Text != string.Empty)
            {
                s = s.Where(q => q.郵便番号2.Contains(txtsZip2.Text)).OrderBy(q => q.登録番号);
            }

            // 住所
            if (txtsAdd.Text != string.Empty)
            {
                s = s.Where(q => (q.住所1).Contains(Utility.getStrConv(txtsAdd.Text))).OrderBy(q => q.登録番号);
            }

            // フリガナ氏名
            if (txtsFuri.Text != string.Empty)
            {
                s = s.Where(q => q.氏名.Contains(Utility.getStrConv(txtsFuri.Text))).OrderBy(q => q.登録番号);
            }

            // TEL/携帯
            if (txtsTel1.Text != string.Empty)
            {
                s = s.Where(q => q.TEL携帯.Contains(txtsTel1.Text)).OrderBy(q => q.登録番号);
            }

            if (txtsTel2.Text != string.Empty)
            {
                s = s.Where(q => q.TEL携帯2.Contains(txtsTel2.Text)).OrderBy(q => q.登録番号);
            }

            if (txtsTel3.Text != string.Empty)
            {
                s = s.Where(q => q.TEL携帯3.Contains(txtsTel3.Text)).OrderBy(q => q.登録番号);
            }

            //// 県警用CSV作成：作成済み
            //if (comboBox1.SelectedIndex == 1 && !dateTimePicker1.Checked)
            //{
            //    s = s.Where(q => !q.IsCSV作成日Null() && q.CSV作成日 != string.Empty).OrderBy(q => q.登録番号);
            //}

            //// 県警用CSV作成：作成日付指定
            //if (comboBox1.SelectedIndex == 1 && dateTimePicker1.Checked)
            //{
            //    string dt = dateTimePicker1.Value.ToShortDateString().Replace("/", "");
            //    s = s.Where(q => !q.IsCSV作成日Null() && q.CSV作成日.Contains(dt)).OrderBy(q => q.登録番号);
            //}

            //// 県警用CSV作成：未作成
            //if (comboBox1.SelectedIndex == 2)
            //{
            //    s = s.Where(q => q.IsCSV作成日Null() || q.CSV作成日 == string.Empty).OrderBy(q => q.登録番号);
            //}

            //// 除外データ
            //if (chkJyogai.Checked)
            //{
            //    s = s.Where(q => !q.Is除外Null() && q.除外 == global.flgOn).OrderBy(q => q.登録番号);
            //}
            //else
            //{
            //    s = s.Where(q => q.Is除外Null() || q.除外 == global.flgOff).OrderBy(q => q.登録番号);
            //}

            // ラベル名
            if (txtLabel.Text != string.Empty)
            {
                s = s.Where(q => q.ラベル.Contains(txtLabel.Text)).OrderBy(q => q.登録番号);
            }

            // 処理担当者名
            if (txtName.Text != string.Empty)
            {
                s = s.Where(q => q.処理担当者.Contains(txtName.Text)).OrderBy(q => q.登録番号);
            }

            int iX = 0;
            
            // dg.Rows.Clear(); // 2019/06/25 コメント化

            // 2019/06/25
            System.Threading.Thread.Sleep(100);
            Application.DoEvents();

            if (s.Count() > 0)
            {
                dg.Rows.Add(s.Count());
            }

            foreach (var t in s)
            {
                //dg.Rows.Add();

                if (t.データ区分 == global.flgOff)
                {
                    dg[coldKbn, iX].Value = "自転車";
                }
                else
                {
                    dg[coldKbn, iX].Value = "原付";
                }

                dg[colCPA, iX].Value = t.登録番号;
                dg[colCarbodyNum, iX].Value = t.車体番号;
                dg[colyymmdd, iX].Value = "20" + t.登録年 + "/" + t.登録月.PadLeft(2, '0') + "/" + t.登録日.PadLeft(2, '0');
                dg[colMaker, iX].Value = t.メーカー;
                dg[colColor, iX].Value = t.塗色;
                dg[colCarStyle, iX].Value = getCarStyleName(t.車種.ToString().PadLeft(2, '0'));

                if (t.Is車両番号1Null())
                {
                    dg[colSharyoNum, iX].Value = t.車両番号2;
                }
                else
                {
                    dg[colSharyoNum, iX].Value = t.車両番号1 + t.車両番号2;
                }

                dg[colCarName, iX].Value = t.車名;
                dg[colZip, iX].Value = t.郵便番号1 + "-" + t.郵便番号2;
                dg[colAdd, iX].Value = t.住所1.Trim();
                dg[colFuri, iX].Value = t.氏名;
                dg[colTel, iX].Value = t.TEL携帯.Trim() + "-" + t.TEL携帯2.Trim() + "-" + t.TEL携帯3.Trim();

                //if (t.IsCSV作成日Null())
                //{
                //    dg[colCsv, iX].Value = string.Empty;
                //}
                //else
                //{
                //    dg[colCsv, iX].Value = t.CSV作成日;
                //}

                dg[colID, iX].Value = t.ID;

                //if (!t.Is除外Null() && t.除外 == global.flgOn)
                //{
                //    dg[colJyogai, iX].Value = "◯";
                //}
                //else
                //{
                //    dg[colJyogai, iX].Value = "";
                //}

                dg[colLabel, iX].Value = t.ラベル;
                dg[colName, iX].Value = t.処理担当者;

                iX++;
            }

            if (s.Count() > 0)
            {
                dg.CurrentCell = null;
                linkLabel2.Enabled = true;

                // 2019/11/15
                label22.Text = "該当件数：" + s.Count().ToString("#,##0") + "件";
            }
            else
            {
                // 2019/06/25
                this.Cursor = Cursors.Default; 
                MessageBox.Show("条件に該当するデータはありませんでした", "検索結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                linkLabel2.Enabled = false;

                // 2019/11/15
                label22.Text = "該当件数： 0件";
            }

            System.Threading.Thread.Sleep(500);
            Application.DoEvents();

            // 2019/06/25
            this.Cursor = Cursors.Default; 
            return s.Count();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void frmCardList_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     車種配列から車種名を取得します </summary>
        /// <param name="cCode">
        ///     車種コード</param>
        /// <returns>
        ///     車種名</returns>
        ///------------------------------------------------------------
        private string getCarStyleName(string cCode)
        {
            string cN = string.Empty;

            for (int i = 0; i < g.arrStyle.GetLength(0); i++)
            {
                if (g.arrStyle[i, 0] == cCode)
                {
                    cN = g.arrStyle[i, 1];
                    break;
                }
            }

            return cN;
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MyLibrary.CsvOut.GridView(dg, "スキャンデータ");
        }

        private void dg_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // 指定データとカード画像を表示
                int iX = Utility.StrtoInt(dg[colID, e.RowIndex].Value.ToString());
                showScanData(iX);

                //// データ再表示
                //dataShow();
            }
        }

        private void showScanData(int iX)
        {
            this.Hide();
            frmScanData frm = new frmScanData(iX);
            frm.ShowDialog();
            this.Show();
        }

        private void txtsZip1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        private void FrmScanList_Shown(object sender, EventArgs e)
        {
            txtsYY.Text = (DateTime.Now.Year - 2000).ToString();
        }
    }
}
