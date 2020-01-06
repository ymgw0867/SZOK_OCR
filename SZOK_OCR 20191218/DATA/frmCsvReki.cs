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
    public partial class frmCsvReki : Form
    {
        public frmCsvReki()
        {
            InitializeComponent();
        }

        cardDataSet dts = new cardDataSet();
        cardDataSetTableAdapters.CSV作成履歴TableAdapter adp = new cardDataSetTableAdapters.CSV作成履歴TableAdapter();

        // カラム定義
        string colDate = "col1";
        string colTekiyou = "col2";
        string colCnt = "col3";
        string colPc = "col4";

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
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 382;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlLight;

                //各列幅指定
                tempDGV.Columns.Add(colDate, "作成日時");
                tempDGV.Columns.Add(colTekiyou, "カード種別");
                tempDGV.Columns.Add(colCnt, "作成件数");
                tempDGV.Columns.Add(colPc, "ＰＣ名");

                tempDGV.Columns[colDate].Width = 150;
                tempDGV.Columns[colTekiyou].Width = 120;
                tempDGV.Columns[colCnt].Width = 110;
                tempDGV.Columns[colPc].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[colDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colCnt].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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

        private void frmCsvReki_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッド定義
            GridViewSetting(dg);

            adp.Fill(dts.CSV作成履歴);
            dataShow(dg);
        }

        private void dataShow(DataGridView gv)
        {
            int iX = 0;

            foreach (var t in dts.CSV作成履歴.OrderByDescending(a => a.作成年月日))
            {
                gv.Rows.Add();

                if (t.Is作成年月日Null())
                {
                    gv[colDate, iX].Value = string.Empty;
                }
                else
                {
                    gv[colDate, iX].Value = t.作成年月日;
                }

                if (t.Is摘要Null())
                {
                    gv[colTekiyou, iX].Value = string.Empty;
                }
                else
                {
                    gv[colTekiyou, iX].Value = t.摘要;
                }

                if (t.Is出力件数Null())
                {
                    gv[colCnt, iX].Value = string.Empty;
                }
                else
                {
                    gv[colCnt, iX].Value = t.出力件数.ToString("#,##0");
                }

                if (t.IsPC名Null())
                {
                    gv[colPc, iX].Value = string.Empty;
                }
                else
                {
                    gv[colPc, iX].Value = t.PC名;
                }
                
                iX++;
            }

            if (gv.RowCount > 0)
            {
                gv.CurrentCell = null;
            }
        }

        private void frmCsvReki_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }
    }
}
