using DocumentFormat.OpenXml.Office.CustomUI;
using MyLibrary;
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
    public partial class frmEraChangeList : Form
    {
        public frmEraChangeList()
        {
            InitializeComponent();
        }

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
        string colAddKN = "col12";
        string colFuri = "col13";
        string colTel = "col14";
        string colID = "colID";
        string colCsv = "col15";
        string colJyogai = "col16";
        string colEraDate = "col17";

        // データグリッドビューカラム定義：変更届
        string colUpdateday = "col18";
        string colHd1 = "col19";
        string colHd2 = "col20";

        bool EditStatus = false;      // 2026/09/07

        private void frmEraChangeList_Load(object sender, EventArgs e)
        {
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッドビュー定義
            gridViewSet(dgChange);  // 変更届
            gridViewSetting(dg);    // 抹消

            // 検索欄初期化
            dispInitial();

            btnExcel.Enabled = false;
            btnClose.Enabled = true;
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        ///-----------------------------------------------------------------
        private void dispInitial()
        {
            txtsYY.Text = string.Empty;
            txtsMM.Text = string.Empty;
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
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Yu Gothic UI", 10, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Yu Gothic UI", 10, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 362;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(216, 232, 254); 
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;                

                //各列幅指定
                tempDGV.Columns.Add(colEraDate, "抹消日");
                tempDGV.Columns.Add(coldKbn, "種別");
                tempDGV.Columns.Add(colCPA, "登録番号");
                tempDGV.Columns.Add(colCarbodyNum, "車体番号");
                tempDGV.Columns.Add(colyymmdd, "登録年月日");
                tempDGV.Columns.Add(colMaker, "メーカー");
                tempDGV.Columns.Add(colColor, "カラー");
                tempDGV.Columns.Add(colCarStyle, "車種");
                tempDGV.Columns.Add(colZip, "〒");
                tempDGV.Columns.Add(colAddKN, "住所");
                tempDGV.Columns.Add(colAdd, "住所フリガナ");
                tempDGV.Columns.Add(colFuri, "氏名");
                tempDGV.Columns.Add(colTel, "ＴＥＬ／携帯");
                tempDGV.Columns.Add(colSharyoNum, "車両番号");
                tempDGV.Columns.Add(colCarName, "車名");
                tempDGV.Columns.Add(colCsv, "静岡県警用CSV作成");
                tempDGV.Columns.Add(colJyogai, "除外");
                tempDGV.Columns.Add(colID, "");

                tempDGV.Columns[colID].Visible = false;

                tempDGV.Columns[colEraDate].Width = 110;
                tempDGV.Columns[coldKbn].Width = 90;
                tempDGV.Columns[colCPA].Width = 140;
                tempDGV.Columns[colCarbodyNum].Width = 200;
                tempDGV.Columns[colyymmdd].Width = 110;
                tempDGV.Columns[colMaker].Width = 100;
                tempDGV.Columns[colColor].Width = 100;
                tempDGV.Columns[colCarStyle].Width = 100;
                tempDGV.Columns[colSharyoNum].Width = 70;
                tempDGV.Columns[colCarName].Width = 70;
                tempDGV.Columns[colZip].Width = 100;
                tempDGV.Columns[colAddKN].Width = 220;
                tempDGV.Columns[colAdd].Width = 330;
                tempDGV.Columns[colFuri].Width = 150;
                tempDGV.Columns[colTel].Width = 140;
                tempDGV.Columns[colCsv].Width = 170;
                tempDGV.Columns[colJyogai].Width = 60;

                tempDGV.Columns[colEraDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colyymmdd].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colZip].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colCsv].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colJyogai].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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

        private void gridViewSet(DataGridView dg)
        {
            dg.Columns.Clear();
            dg.Rows.Clear();

            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                dg.EnableHeadersVisualStyles = false;
                dg.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.SteelBlue;
                dg.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;

                // 列ヘッダー表示位置指定
                dg.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 列ヘッダーフォント指定
                //dg.ColumnHeadersDefaultCellStyle.Font = new Font("游明朝", 9, FontStyle.Regular);
                dg.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Yu Gothic UI", 10, FontStyle.Regular);

                // データフォント指定
                //dg.DefaultCellStyle.Font = new Font("游明朝", 10, FontStyle.Regular);
                dg.DefaultCellStyle.Font = new System.Drawing.Font("Yu Gothic UI", 10, FontStyle.Regular);

                // 行の高さ
                dg.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                dg.ColumnHeadersHeight = 22;
                dg.RowTemplate.Height = 22;

                // 全体の高さ
                dg.Height = 398;

                // 奇数行の色
                ////dg.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.Control;
                //dg.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(216, 232, 254);
                ////dg.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;

                dg.Columns.Add(colUpdateday, "変更日");
                dg.Columns.Add(colCPA, "登録番号");
                dg.Columns.Add(colHd1, "");
                dg.Columns.Add(colZip, "〒");
                dg.Columns.Add(colAddKN, "住所");
                dg.Columns.Add(colAdd, "住所フリガナ");
                dg.Columns.Add(colFuri, "氏名");
                dg.Columns.Add(colTel, "ＴＥＬ／携帯");
                dg.Columns.Add(colHd2, "");

                dg.Columns[colUpdateday].Width = 110;
                dg.Columns[colCPA].Width = 140;
                dg.Columns[colHd1].Width = 60;
                dg.Columns[colZip].Width = 100;
                dg.Columns[colAddKN].Width = 220;
                dg.Columns[colAdd].Width = 330;
                dg.Columns[colFuri].Width = 150;
                dg.Columns[colTel].Width = 140;
                dg.Columns[colHd2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                dg.Columns[colUpdateday].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dg.Columns[colCPA].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dg.Columns[colZip].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dg.Columns[colAddKN].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                dg.Columns[colAdd].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                dg.Columns[colFuri].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                dg.Columns[colTel].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;

                // 編集可否
                dg.ReadOnly = true;

                // 行ヘッダを表示しない
                dg.RowHeadersVisible = false;

                // 選択モード
                dg.SelectionMode = DataGridViewSelectionMode.CellSelect;
                dg.MultiSelect = false;

                // 追加行表示しない
                dg.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                dg.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                dg.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                dg.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                dg.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //dg.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                // 罫線
                dg.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                dg.CellBorderStyle = DataGridViewCellBorderStyle.None;

                //dg.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
                //dg.CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical;
                //dg.GridColor = System.Drawing.Color.SteelBlue;

                // 列固定
                //dg.Columns[colTimes].Frozen = true;
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

            if (Utility.StrtoInt(txtsMM.Text) == global.flgOff)
            {
                MessageBox.Show("登録月を必ず指定してください", "検索項目指定", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (Utility.StrtoInt(txtsMM.Text) < 1 || Utility.StrtoInt(txtsMM.Text) > 12)
            {
                MessageBox.Show("登録月は1～12の範囲で指定してください", "検索項目指定", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // データグリッドに変更届データを表示
            DataFindUpdate();

            // データグリッドに抹消データを表示
            DataFindErasure();
        }


        private int DataFindUpdate()
        {
            label3.Text = string.Empty;

            dgChange.Rows.Clear();

            System.Threading.Thread.Sleep(10);
            Application.DoEvents();

            this.Cursor = Cursors.WaitCursor;

            // SQL Server接続
            var master = new ClsMaster(Properties.Settings.Default.sServerName, Properties.Settings.Default.sLogin,
                                   Properties.Settings.Default.sPass, Properties.Settings.Default.sDatabase);

            var result = master.ReadUpdateData(int.Parse(txtsYY.Text), int.Parse(txtsMM.Text));


            int iX = 0;

            System.Threading.Thread.Sleep(10);
            Application.DoEvents();

            if (result.Count() > 0)
            {
                dgChange.Rows.Add(result.Count() * 2);
            }

            foreach (var t in result)
            {
                dgChange[colUpdateday, iX].Value = t.UpdateDay.ToString("yyyy/MM/dd");
                dgChange[colCPA, iX].Value = t.Number;

                dgChange[colHd1, iX].Value = "変更前";
                dgChange[colZip, iX].Value = t.OldZipCode1 != string.Empty ? t.OldZipCode1 + "-" + t.OldZipCode2 : string.Empty;
                dgChange[colAddKN, iX].Value = t.OldAddressKanji.Trim();
                dgChange[colAdd, iX].Value = t.OldAddress1.Trim();
                dgChange[colFuri, iX].Value = t.OldName;
                dgChange[colTel, iX].Value = t.OldMobile1 != string.Empty ? t.OldMobile1.Trim() + "-" + t.OldMobile2.Trim() + "-" + t.OldMobile3.Trim() : string.Empty;
                iX++;
                dgChange[colHd1, iX].Value = "変更後";
                dgChange[colZip, iX].Value = t.NewZipCode1 != string.Empty ? t.NewZipCode1 + "-" + t.NewZipCode2 : string.Empty;
                dgChange[colAddKN, iX].Value = t.NewAddressKanji.Trim();
                dgChange[colAdd, iX].Value = t.NewAddress1.Trim();
                dgChange[colFuri, iX].Value = t.NewName;
                dgChange[colTel, iX].Value = t.NewMobile1 != string.Empty ? t.NewMobile1.Trim() + "-" + t.NewMobile2.Trim() + "-" + t.NewMobile3.Trim() : string.Empty;
             

                if ((iX + 1) % 4 == 0)
                {
                    dgChange.Rows[iX - 1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(216, 232, 254);
                    dgChange.Rows[iX].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(216, 232, 254);
                }

                iX++;
            }

            if (result.Count() > 0)
            {
                dgChange.CurrentCell = null;
                btnExcel.Enabled = true;

                // 2019/11/15
                label3.Text = "該当件数：" + result.Count().ToString("#,##0") + "件";
            }
            else
            {
                // 2019/06/25
                this.Cursor = Cursors.Default;
                MessageBox.Show("条件に該当するデータはありませんでした", "検索結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnExcel.Enabled = false;

                btnExcel.Enabled = false;
                btnClose.Enabled = false;

                // 2019/11/15
                label3.Text = "該当件数： 0件";
            }

            System.Threading.Thread.Sleep(500);
            Application.DoEvents();

            // 2019/06/25
            this.Cursor = Cursors.Default;
            return result.Count();
        }


        /// <summary>
        ///    データグリッドに抹消データを表示する
        /// </summary>
        /// <returns>表示したデータの件数</returns>
        private int DataFindErasure()
        {
            label22.Text = string.Empty;    // 2019/11/15

            dg.Rows.Clear();

            System.Threading.Thread.Sleep(10);
            Application.DoEvents();

            this.Cursor = Cursors.WaitCursor;

            // SQL Server接続
            var master = new ClsMaster(Properties.Settings.Default.sServerName, Properties.Settings.Default.sLogin,
                                   Properties.Settings.Default.sPass, Properties.Settings.Default.sDatabase);

            var result = master.ReadErasureData(int.Parse(txtsYY.Text), int.Parse(txtsMM.Text));


            int iX = 0;

            // 2019/06/25
            System.Threading.Thread.Sleep(10);
            Application.DoEvents();

            if (result.Count() > 0)
            {
                dg.Rows.Add(result.Count());
            }

            foreach (var t in result)
            {
                dg[colEraDate, iX].Value = t.UpDate.ToString("yyyy/MM/dd");

                if (t.DataCategory == global.flgOff)
                {
                    dg[coldKbn, iX].Value = "自転車";
                }
                else
                {
                    dg[coldKbn, iX].Value = "原付";
                }

                dg[colCPA, iX].Value = t.Number;
                dg[colCarbodyNum, iX].Value = t.VehicleIdentificationNumber;
                dg[colyymmdd, iX].Value = "20" + t.AddYear + "/" + t.AddMonth.PadLeft(2, '0') + "/" + t.AddDay.PadLeft(2, '0');
                dg[colMaker, iX].Value = t.Maker;
                dg[colColor, iX].Value = t.Color;
                dg[colCarStyle, iX].Value = getCarStyleName(t.CarModel.ToString().PadLeft(2, '0'));
                dg[colZip, iX].Value = t.ZipCode1 + "-" + t.ZipCode2;
                dg[colAddKN, iX].Value = t.AddressKanji.Trim();
                dg[colAdd, iX].Value = t.Address1.Trim();
                dg[colFuri, iX].Value = t.Name;
                dg[colTel, iX].Value = t.Mobile1.Trim() + "-" + t.Mobile2.Trim() + "-" + t.Mobile3.Trim();
                dg[colID, iX].Value = t.ID;
                dg[colCsv, iX].Value = t.CsvCreationDate;

                if (t.Exception == global.flgOn)
                {
                    dg[colJyogai, iX].Value = "◯";
                }
                else
                {
                    dg[colJyogai, iX].Value = "";
                }

                dg[colSharyoNum, iX].Value = t.VehicleNumber1 + t.VehicleNumber2;
                dg[colCarName, iX].Value = t.CarName;

                iX++;
            }

            if (result.Count() > 0)
            {
                dg.CurrentCell = null;
                btnExcel.Enabled = true;

                // 2019/11/15
                label22.Text = "該当件数：" + result.Count().ToString("#,##0") + "件";
            }
            else
            {
                // 2019/06/25
                this.Cursor = Cursors.Default;
                MessageBox.Show("条件に該当するデータはありませんでした", "検索結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnExcel.Enabled = false;
                
                // 2019/11/15
                label22.Text = "該当件数： 0件";
            }

            System.Threading.Thread.Sleep(500);
            Application.DoEvents();

            // 2019/06/25
            this.Cursor = Cursors.Default;
            return result.Count();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void frmEraChangeList_FormClosing(object sender, FormClosingEventArgs e)
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
            MyLibrary.CsvOut.GridView(dgChange, "防犯登録カードデータ");
        }


        private void txtsZip1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        private void FrmEraChangeList_Shown(object sender, EventArgs e)
        {
            txtsYY.Text = (DateTime.Now.Year - 2000).ToString();
            txtsMM.Focus();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void dg_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
