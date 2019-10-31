using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using JS_OCR.Common;

namespace JS_OCR
{
    public partial class frmComSelect : Form
    {
        public frmComSelect()
        {
            InitializeComponent();

            // 選択会社情報初期化
            _pblComNo = string.Empty;       // 会社№
            _pblComName = string.Empty;     // 会社名
            _pblDbName = string.Empty;      // データベース名
        }

        #region 選択会社取得情報プロパティ
        public string _pblComNo { get; private set; }       // 会社№
        public string _pblComName { get; private set; }     // 会社名
        public string _pblDbName { get; private set; }      // 会社データベース名
        #endregion

        #region データグリッドビューカラム定義
        const string C_1 = "col1";
        const string C_2 = "col2";
        const string C_3 = "col3";
        const string C_4 = "col4";
        const string C_5 = "col5";
        #endregion

        private void frmComSelect_Load(object sender, EventArgs e)
        {
            // ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // ウィンドウズ最大サイズ
            Utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            // DataGridViewの設定
            GridViewSetting(dg1);

            // データ表示
            GridViewShowData(dg1);

            // 終了時タグ初期化
            Tag = string.Empty;
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います</summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        /// ------------------------------------------------------------------------------
        public void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                // フォームサイズ定義

                // 列スタイルを変更する
                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9.5, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 18;
                tempDGV.RowTemplate.Height = 18;

                // 全体の高さ
                tempDGV.Height = 180;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                tempDGV.Columns.Add(C_1, "No");
                tempDGV.Columns.Add(C_2, "会社名");
                tempDGV.Columns.Add(C_4, "処理年度");
                tempDGV.Columns.Add(C_3, "DCODE");

                tempDGV.Columns[C_3].Visible = false; //データベース名は非表示

                tempDGV.Columns[C_1].Width = 100;
                tempDGV.Columns[C_2].Width = 300;
                tempDGV.Columns[C_4].Width = 100;

                tempDGV.Columns[C_2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[C_1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[C_2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                tempDGV.Columns[C_4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

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

                // 列サイズ変更禁止
                tempDGV.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// -------------------------------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ会社情報を表示する</summary>
        /// <param name="tempDGV">
        ///     DataGridViewオブジェクト名</param>
        /// -------------------------------------------------------------------------------------
        private void GridViewShowData(DataGridView tempDGV)
        {
            string sqlSTRING = string.Empty;

            dbControl.DataControl sdcon = new dbControl.DataControl(Properties.Settings.Default.SQLDataBase);
            SqlDataReader dR;

            // データリーダーを取得する
            sqlSTRING += "SELECT KCODE,DCODE,CONAME1,STDATE FROM SELDATA ";
            sqlSTRING += "order by KCODE, STDATE desc";

            dR = sdcon.FreeReader(sqlSTRING);

            try
            {
                // グリッドビューに表示する
                int iX = 0;
                tempDGV.RowCount = 0;

                while (dR.Read())
                {
                    // データグリッドにデータを表示する
                    tempDGV.Rows.Add();
                    GridViewCellData(tempDGV, iX, dR);
                    iX++;
                }
                tempDGV.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                dR.Close();
                sdcon.Close();
            }

            // 会社情報がないとき
            if (tempDGV.RowCount == 0)
            {
                MessageBox.Show("会社情報が存在しません", "会社選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                Environment.Exit(0);
            }
        }

        /// <summary>
        /// データグリッドに表示データをセットする
        /// </summary>
        /// <param name="tempDGV">datagridviewオブジェクト名</param>
        /// <param name="iX">Row№</param>
        /// <param name="dR">データリーダーオブジェクト名</param>
        private void GridViewCellData(DataGridView tempDGV, int iX, SqlDataReader dR)
        {
            tempDGV[C_1, iX].Value = dR["KCODE"].ToString();            // 会社№
            tempDGV[C_2, iX].Value = dR["CONAME1"].ToString().Trim();   // 会社名
            tempDGV[C_3, iX].Value = dR["DCODE"].ToString().Trim();     // DCODE（非表示項目）

            // 処理年度 2014/03/28
            DateTime dt;
            if (DateTime.TryParse(dR["STDATE"].ToString(), out dt))
            {
                int waYear = dt.Year - Properties.Settings.Default.RekiHosei;
                tempDGV[C_4, iX].Value = Properties.Settings.Default.gengou + waYear.ToString() + "年";
            }
            else
            {
                tempDGV[C_4, iX].Value = "不明";
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            // 会社情報がないときはそのままクローズ
            if (dg1.RowCount == 0)
            {
                _pblComNo = string.Empty;       // 会社№
                _pblDbName = string.Empty;      // データベース名
            }
            else
            {
                if (dg1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("会社を選択してください", "会社未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                // 選択した会社情報を取得する
                _pblComNo = dg1[C_1, dg1.SelectedRows[0].Index].Value.ToString();     // 会社№
                _pblComName = dg1[C_2, dg1.SelectedRows[0].Index].Value.ToString();   // 会社名
                _pblDbName = dg1[C_3, dg1.SelectedRows[0].Index].Value.ToString();    // データベース名
            }

            // フォームを閉じる
            Tag = "btn";
            this.Close();
        }

        private void frmComSelect_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                if (Tag.ToString() == string.Empty)
                {
                    if (MessageBox.Show("プログラムを終了します。よろしいですか？", "終了", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        // 終了処理
                        Environment.Exit(0);
                    }
                    else
                    {
                        e.Cancel = true;
                        return;
                    }
                }
            }

            //this.Dispose();
        }    
    }
}
