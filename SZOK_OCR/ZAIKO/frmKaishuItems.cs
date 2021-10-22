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
    public partial class frmKaishuItems : Form
    {
        public frmKaishuItems(ClsShukko shukko)
        {
            InitializeComponent();

            clsShukko = shukko; 
        }

        cardDataSet dts = new cardDataSet();
        cardDataSetTableAdapters.回収データTableAdapter kAdp = new cardDataSetTableAdapters.回収データTableAdapter();

        ClsShukko clsShukko = new ClsShukko();

        // カラム定義
        private readonly string colSeq    = "c0";
        private readonly string colNum    = "c1";
        private readonly string colStatus = "c2";
        private readonly string colDate   = "c3";
        private readonly string colID     = "c4";

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
            dg1.CurrentCell = null;
        }

        private void KaishuDataShow()
        {
            // 行数を設定して表示色を初期化
            dg1.Rows.Clear();
            dg1.Rows.Add(clsShukko.Suu);
            
            int seqNum = 1;
            for (int i = clsShukko.StartNumber; i <= clsShukko.EndNumber; i++)
            {
                dg1[colSeq,    seqNum - 1].Value = seqNum;
                dg1[colNum,    seqNum - 1].Value = i;

                dg1[colStatus, seqNum - 1].Value = "";
                dg1[colDate,   seqNum - 1].Value = "";
                dg1[colID,     seqNum - 1].Value = "";

                foreach (var t in dts.回収データ.Where(a => a.登録番号 == i))
                {
                    dg1[colStatus, seqNum - 1].Value = "○";
                    dg1[colDate,   seqNum - 1].Value = t.回収年月日.ToShortDateString();
                    dg1[colID,     seqNum - 1].Value = t.ID;
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

            //カレントセル選択状態としない
            dg1.CurrentCell = null;
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
                tempDGV.Columns.Add(colSeq,    "");
                tempDGV.Columns.Add(colNum,    "登録番号");
                tempDGV.Columns.Add(colStatus, "回収");
                tempDGV.Columns.Add(colDate,   "回収日");
                tempDGV.Columns.Add(colID,      "");

                // 各列幅指定
                tempDGV.Columns[colSeq].Width    = 50;
                tempDGV.Columns[colNum].Width    = 100;
                tempDGV.Columns[colStatus].Width = 110;
                //tempDGV.Columns[colDate].Width   = 165;

                tempDGV.Columns[colID].Visible = false;
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
                tempDGV.MultiSelect = false;

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
        }
    }
}
