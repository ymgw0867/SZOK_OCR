using System;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using System.Data.SqlClient;
using System.Text;
using SZOK_OCR.Common;

namespace SZOK_OCR.OCR
{
    public partial class frmSelectLabel : Form
    {
        public frmSelectLabel()
        {
            InitializeComponent();

            adp.Fill(dtsC.SCAN_DATA);
            dAdp.Fill(dts.防犯カード);
        }

        cardDataSet dtsC = new cardDataSet();
        cardDataSetTableAdapters.SCAN_DATATableAdapter adp = new cardDataSetTableAdapters.SCAN_DATATableAdapter();

        szokDataSet dts = new szokDataSet();
        szokDataSetTableAdapters.防犯カードTableAdapter dAdp = new szokDataSetTableAdapters.防犯カードTableAdapter();

        string colChk = "c0";
        string colLabel = "c1";
        string colName = "c2";
        string colCount = "c3";
        string colID = "c4";

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '\t')
            {
                e.Handled = true;
            }
        }

        private void frmFaxSelect_Load(object sender, EventArgs e)
        {
            // フォーム最大値
            MaximumSize = new System.Drawing.Size(Width, Height);

            // フォーム最小値
            MinimumSize = new System.Drawing.Size(Width, Height);

            // DataGridViewの設定
            GridViewSetting(Dgv1);

            // ラベル毎のスキャンデータ件数表示
            getScanData(Dgv1);

            // 処理中データ件数表示
            int val = GetLocalData();

            if (val > 0)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }

        private int GetLocalData()
        {
            int val = dts.防犯カード.Count();
            lblDataCnt.Text = val.ToString("#,##0");

            // 前バージョンのＯＣＲ認識後データ
            int ocrCnt = System.IO.Directory.GetFiles(Properties.Settings.Default.dataPath, "*.csv").Count();
            lblOCRCnt.Text = ocrCnt.ToString("#,##0");

            return val + ocrCnt;
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     ラベル毎の件数をグリッドビューに表示 </summary>
        /// <param name="dg">
        ///     データグリッドビュー</param>
        ///-----------------------------------------------------------------
        private void getScanData(DataGridView dg)
        {
            if (dtsC.SCAN_DATA.Count == 0)
            {
                // データがなければ終了
                return;
            }

            var ss = dtsC.SCAN_DATA.Select(a => a.ラベル).Distinct();

            dg.Rows.Add(ss.Count());

            int iX = 0;

            // データグリッドビューに表示
            foreach (var t in ss)
            {
                dg[colChk, iX].Value = "false";
                dg[colLabel, iX].Value = t;
                dg[colName, iX].Value = dtsC.SCAN_DATA.Where(a => a.ラベル == t).Select(a => a.処理担当者).First();
                dg[colCount, iX].Value = dtsC.SCAN_DATA.Where(a => a.ラベル == t).Count();

                iX++;
            }
            
            dg.CurrentCell = null;
        }
        
        ///-----------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///-----------------------------------------------------------------------
        public void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する
                tempDGV.EnableHeadersVisualStyles = false;
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", (float)10, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 402;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列幅指定
                DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
                chk.Name = colChk;
                tempDGV.Columns.Add(chk);
                tempDGV.Columns[colChk].HeaderText = "";

                tempDGV.Columns.Add(colLabel, "ラベル");
                tempDGV.Columns.Add(colName, "処理担当者");
                tempDGV.Columns.Add(colCount, "ＯＣＲ認識件数");

                tempDGV.Columns[colChk].Width = 40;
                tempDGV.Columns[colLabel].Width = 120;
                //tempDGV.Columns[colName].Width = 300;
                tempDGV.Columns[colCount].Width = 120;

                tempDGV.Columns[colLabel].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colCount].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[colName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                // 編集可否
                tempDGV.ReadOnly = false;

                // 編集可否
                foreach (DataGridViewColumn item in tempDGV.Columns)
                {
                    // チェックボックスのみ使用可
                    if (item.Name == colChk)
                    {
                        tempDGV.Columns[item.Name].ReadOnly = false;
                    }
                    else
                    {
                        tempDGV.Columns[item.Name].ReadOnly = true;
                    }
                }

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

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

                //// 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;

                ////ソート機能制限
                //for (int i = 0; i < tempDGV.Columns.Count; i++)
                //{
                //    tempDGV.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //}

                // 罫線
                tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MyCnt = 0;
            MyBool = false;
            Close();
        }

        public int MyCnt { get; set; }
        public bool MyBool { get; set; }

        private void button1_Click(object sender, EventArgs e)
        {
            int n = getDataCount();

            if (n == 0)
            {
                MessageBox.Show("処理するデータがありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else
            {
                if (MessageBox.Show("防犯登録データ作成を開始します。よろしいですか。", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }

            MyBool = true;

            // OCR認識したスキャンデータを取り込む
            if (n > 0)
            {
                GetScanToData();
            }

            Close();
        }
        
        ///-----------------------------------------------------------
        /// <summary>
        ///     処理データ件数を取得する </summary>
        /// <returns>
        ///     データ件数 </returns>
        ///-----------------------------------------------------------
        private int getDataCount()
        {
            int dCnt = 0;

            for (int i = 0; i < Dgv1.Rows.Count; i++)
            {
                if (Dgv1[colChk, i].Value.ToString() == "True")
                {
                    dCnt += Utility.StrtoInt(Utility.NulltoStr(Dgv1[colCount, i].Value));
                }
            }

            dCnt += Utility.StrtoInt(lblOCRCnt.Text) + Utility.StrtoInt(lblDataCnt.Text);

            return dCnt;
        }

        ///-------------------------------------------------------
        /// <summary>
        ///     スキャンデータを自分の防犯登録データへ取り込む </summary>
        /// <param name="mCnt">
        ///     取り込む枚数</param>
        ///-------------------------------------------------------
        private void GetScanToData()
        {
            for (int i = 0; i < Dgv1.Rows.Count; i++)
            {
                if (Dgv1[colChk, i].Value.ToString() == "True")
                {
                    string lbl = Utility.nulltoStr2(Dgv1[colLabel, i].Value);

                    adp.FillByLabel(dtsC.SCAN_DATA, lbl);

                    foreach (var t in dtsC.SCAN_DATA)
                    {
                        // 共有のSCAN_DATAをローカルの防犯登録データに出力
                        dAdp.Insert(t.データ区分, t.画像名, t.登録年, t.登録月, t.登録日, t.登録番号, t.車体番号,
                            t.メーカー, t.塗色, t.車種, t.郵便番号1, t.郵便番号2, t.車両番号1, t.車両番号2, t.車名,
                            t.住所漢字, t.住所1, string.Empty, t.氏名, t.TEL携帯, t.TEL携帯2, t.TEL携帯3, t.PC名, t.CSV作成日,
                            t.備考, t.更新年月日, 0);

                        // 画像をローカルのDATAフォルダに移動
                        System.IO.File.Move(Properties.Settings.Default.scanDataPath + t.画像名, Properties.Settings.Default.dataPath + t.画像名);
                    }

                    // 該当スキャンデータ削除
                    adp.DeleteQueryLabel(lbl);
                }
            }
        }

        private void Dgv1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (Dgv1.Rows.Count < 1)
            {
                return;
            }

            if (e.ColumnIndex == 0)
            {
                bool chk = false;

                for (int i = 0; i < Dgv1.Rows.Count; i++)
                {
                    if (Utility.nulltoStr2(Dgv1[colChk, i].Value) == "True")
                    {
                        chk = true;
                        break;
                    }
                }

                if (chk)
                {
                    button1.Enabled = true;
                }
                else
                {
                    if ((Utility.StrtoInt(lblDataCnt.Text) + Utility.StrtoInt(lblOCRCnt.Text)) > global.flgOff)
                    {
                        button1.Enabled = true;
                    }
                    else
                    {
                        button1.Enabled = false;
                    }
                }
            }
        }

        private void Dgv1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (Dgv1.CurrentCellAddress.X == 0)
            {
                if (Dgv1.IsCurrentCellDirty)
                {
                    Dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            }
        }
    }
}
