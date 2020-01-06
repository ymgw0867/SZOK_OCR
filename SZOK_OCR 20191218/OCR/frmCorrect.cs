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
    public partial class frmCorrect : Form
    {
        szokDataSetTableAdapters.防犯カードTableAdapter adp = new szokDataSetTableAdapters.防犯カードTableAdapter();
        szokDataSet dts = new szokDataSet();

        // 検索ＩＤ
        string dID = string.Empty;

        /// <summary>
        ///     カレントデータRowsインデックス</summary>
        int[] cID = null;
        int cI = 0;

        string[] zipArray = null;   // 郵便番号配列

        public frmCorrect(string sID)
        {
            InitializeComponent();

            if (sID == string.Empty)    // ＯＣＲデータを新規登録する
            {
                // 画像パス取得
                global.pblImagePath = Properties.Settings.Default.dataPath;

                // ＯＣＲデータ読み込み   
                adp.Fill(dts.防犯カード);
            }
            else　// 登録済みデータの検索及び編集
            {
                dID = sID;
                adp.Fill(dts.防犯カード);
            }
        }

        private void frmCorrect_Load(object sender, EventArgs e)
        {
            this.pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // 郵便番号CSV配列読み込み
            Utility.zipCsvLoad(ref zipArray);
            
            // 勤務データ登録
            if (dID == string.Empty)
            {
                // CSVデータをMDBへ読み込みます
                GetCsvDataToMDB();

                // データセットへデータを読み込みます
                adp.Fill(dts.防犯カード);
                
                // データテーブル件数カウント
                if (dts.防犯カード.Count == 0)
                {
                    MessageBox.Show("対象となるＯＣＲ防犯登録データがありません", "ＯＣＲ自転車防犯登録データ登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //終了処理
                    Environment.Exit(0);
                }

                // キー配列作成
                keyArrayCreate();
            }

            // キャプション
            this.Text = "ＯＣＲ防犯登録データ登録";

            // 編集作業、過去データ表示の判断
            if (dID == string.Empty) // パラメータのヘッダIDがないときは編集作業
            {
                // 最初のレコードを表示
                cI = 0;
                showOcrData(cI);
            }

            // tagを初期化
            this.Tag = string.Empty;
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     キー配列作成 </summary>
        ///-------------------------------------------------------------
        private void keyArrayCreate()
        {
            int iX = 0;
            foreach (var t in dts.防犯カード.OrderBy(a=> a.ID))
            {
                Array.Resize(ref cID, iX + 1);
                cID[iX] = t.ID;
                iX++;
            }
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVデータをMDBへインサートする</summary>
        ///----------------------------------------------------------------------------
        private void GetCsvDataToMDB()
        {
            // CSVファイルがなければ終了
            if (System.IO.Directory.GetFiles(Properties.Settings.Default.dataPath, "*.csv").Count() == 0)
            {
                return;
            }

            // オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            // OCRのCSVデータをMDBへ取り込む
            OCRData ocr = new OCRData();
            ocr.CsvToMdb(Properties.Settings.Default.dataPath, frmP, dts);

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     伝票画像表示 </summary>
        /// <param name="iX">
        ///     現在の伝票</param>
        /// <param name="tempImgName">
        ///     画像名</param>
        /// ------------------------------------------------------------------------------
        public void ShowImage(string tempImgName)
        {
            //修正画面へ組み入れた画像フォームの表示    
            //画像の出力が無い場合は、画像表示をしない。
            if (tempImgName == string.Empty)
            {
                leadImg.Visible = false;
                lblNoImage.Visible = false;
                global.pblImagePath = string.Empty;
                return;
            }

            //画像ファイルがあるとき表示
            if (System.IO.File.Exists(tempImgName))
            {
                lblNoImage.Visible = false;
                leadImg.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = true;
                btnMinus.Enabled = true;

                //画像ロード
                Leadtools.Codecs.RasterCodecs.Startup();
                Leadtools.Codecs.RasterCodecs cs = new Leadtools.Codecs.RasterCodecs();

                // 描画時に使用される速度、品質、およびスタイルを制御します。 
                Leadtools.RasterPaintProperties prop = new Leadtools.RasterPaintProperties();
                prop = Leadtools.RasterPaintProperties.Default;
                prop.PaintDisplayMode = Leadtools.RasterPaintDisplayModeFlags.Resample;
                leadImg.PaintProperties = prop;

                leadImg.Image = cs.Load(tempImgName, 0, Leadtools.Codecs.CodecsLoadByteOrder.BgrOrGray, 1, 1);

                //画像表示倍率設定
                if (global.miMdlZoomRate == 0f)
                {
                    leadImg.ScaleFactor *= global.ZOOM_RATE;
                }
                else
                {
                    leadImg.ScaleFactor *= global.miMdlZoomRate;
                }

                //画像のマウスによる移動を可能とする
                leadImg.InteractiveMode = Leadtools.WinForms.RasterViewerInteractiveMode.Pan;

                // グレースケールに変換
                Leadtools.ImageProcessing.GrayscaleCommand grayScaleCommand = new Leadtools.ImageProcessing.GrayscaleCommand();
                grayScaleCommand.BitsPerPixel = 8;
                grayScaleCommand.Run(leadImg.Image);
                leadImg.Refresh();

                cs.Dispose();
                Leadtools.Codecs.RasterCodecs.Shutdown();
                //global.pblImagePath = tempImgName;
            }
            else
            {
                //画像ファイルがないとき
                lblNoImage.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = false;
                btnMinus.Enabled = false;

                leadImg.Visible = false;
                //global.pblImagePath = string.Empty;
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if ((cI + 1) < cID.Length)
            {
                // カレントデータ更新
                cuDataUpdate(cID[cI]);

                // 次のレコードを表示
                cI++;
                showOcrData(cI);
            }
        }

        private void btnEnd_Click(object sender, EventArgs e)
        {
            // カレントデータ更新
            cuDataUpdate(cID[cI]);

            // 最後のレコードを表示
            cI = cID.Length - 1;
            showOcrData(cI);
        }

        private void btnBefore_Click(object sender, EventArgs e)
        {
            if (cI > 0)
            {
                // カレントデータ更新
                cuDataUpdate(cID[cI]);

                // 前のレコードを表示
                cI--;
                showOcrData(cI);
            }
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            // カレントデータ更新
            cuDataUpdate(cID[cI]);

            // 最初のレコードを表示
            cI = 0;
            showOcrData(cI);
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
        }

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     防犯カードデータ更新 </summary>
        /// <param name="sID">
        ///     防犯カードデータID</param>
        ///-------------------------------------------------------------------
        private void cuDataUpdate(int sID)
        {
            szokDataSet.防犯カードRow r = dts.防犯カード.Single(a => a.ID == sID);

            if (radioButton1.Checked)
            {
                r.データ区分 = global.flgOff;
            }
            else if (radioButton2.Checked)
            {
                r.データ区分 = global.flgOn;
            }

            r.登録年 = txtYear.Text;
            r.登録月 = txtMonth.Text;
            r.登録日 = txtDay.Text;

            r.登録番号 = kanmaDelete(txtTourokuNum.Text);
            r.車体番号 = kanmaDelete(txtShataiNum.Text);
            r.メーカー = kanmaDelete(txtMaker.Text);
            r.塗色 = kanmaDelete(txtColor.Text);
            r.車種 = Utility.StrtoInt(txtStyle.Text);
            r.郵便番号1 = txtZip1.Text;
            r.郵便番号2 = txtZip2.Text;
            r.車両番号1 = kanmaDelete(txtSharyoNum.Text);
            r.車両番号2 = kanmaDelete(txtSharyoNum2.Text);
            r.車名 = kanmaDelete(txtCarName.Text);
            r.住所漢字 = txtAdd.Text;
            r.住所1 = Utility.GetStringSubMax(kanmaDelete(txtAddFuri.Text), 40);
            r.住所2 = string.Empty;
            r.氏名 = kanmaDelete(txtFuri.Text);
            r.TEL携帯 = txtTel.Text;
            r.TEL携帯2 = txtTel2.Text;
            r.TEL携帯3 = txtTel3.Text;
            r.備考 = txtMemo.Text;
            r.更新年月日 = DateTime.Now;

            // 画面確認チェック 2016/01/25
            if (checkBox1.Checked)
            {
                r.確認 = global.flgOn;
            }
            else
            {
                r.確認 = global.flgOff;
            }
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     シングルコーテーションとカンマを除去 </summary>
        /// <param name="s">
        ///     文字列</param>
        /// <returns>
        ///     変換後文字列</returns>
        ///----------------------------------------------------------------
        private string kanmaDelete(string s)
        {
            return s.Replace("'", "").Replace(",", "");
        }


        private void btnDel_Click(object sender, EventArgs e)
        {
        }

        private void cuDataDelete(int sID)
        {
            szokDataSet.防犯カードRow r = dts.防犯カード.Single(a => a.ID == sID);
            string fName = r.画像名;   // 画像名取得
            r.Delete();     // データ削除

            // データベース更新
            adp.Update(dts.防犯カード);

            // データ再読み込み
            adp.Fill(dts.防犯カード);

            // 画像削除 2016/05/26
            System.IO.File.Delete(Properties.Settings.Default.dataPath + fName);

            // すべてのデータを削除したときは終了します
            if (dts.防犯カード.Count() == 0)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("すべてのＯＣＲ防犯登録カードデータが削除されました。").Append(Environment.NewLine);
                sb.Append("処理を終了します。").Append(Environment.NewLine);
                
                MessageBox.Show(sb.ToString(),"データ削除",MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                return;
            }
            else
            {
                // キー配列再作成
                keyArrayCreate();

                // データ表示
                if (cI > (cID.Length - 1))
                {
                    cI = cID.Length - 1;
                }

                showOcrData(cI);
            }
        }

        private void btnPlus_Click(object sender, EventArgs e)
        {
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > global.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= global.ZOOM_STEP;
            }

            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void leadImg_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Default;
        }

        private void leadImg_MouseMove(object sender, MouseEventArgs e)
        {
            this.Cursor = Cursors.Hand;
        }

        private void btnErrCheck_Click(object sender, EventArgs e)
        {
        }


        /// -----------------------------------------------------------------------------------
        /// <summary>
        ///     エラーチェックを実行する</summary>
        /// <param name="cIdx">
        ///     現在表示中の勤務票ヘッダデータインデックス</param>
        /// <param name="ocr">
        ///     OCRDATAクラスインスタンス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        /// -----------------------------------------------------------------------------------
        private bool getErrData(int cIdx, OCRData ocr)
        {
            // カレントレコード更新
            cuDataUpdate(cID[cI]);

            // エラー番号初期化
            ocr._errNumber = ocr.eNothing;

            // エラーメッセージクリーン
            ocr._errMsg = string.Empty;

            // エラーチェック実行①:カレントレコードから最終レコードまで
            if (!ocr.errCheckMain(cIdx, cID.Length - 1, this, dts, cID, zipArray))
            {
                return false;
            }

            // エラーチェック実行②:最初のレコードからカレントレコードの前のレコードまで
            if (cIdx > 0)
            {
                if (!ocr.errCheckMain(0, (cIdx - 1), this, dts, cID, zipArray))
                {
                    return false;
                }
            }

            // エラーなし
            lblErrMsg.Text = string.Empty;

            return true;
        }

        private void txtStyle_Leave(object sender, EventArgs e)
        {
            txtStyle.Text = txtStyle.Text.PadLeft(2, '0');

            // 車種名取得
            txtStyleName.Text = getStyleName(txtStyle.Text);
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     車種名取得 </summary>
        /// <param name="s">
        ///     車種番号文字列</param>
        /// <returns>
        ///     車種名</returns>
        ///----------------------------------------------------------------
        private string getStyleName(string s)
        {
            global g = new global();
            for (int i = 0; i < g.arrStyle.GetLength(0); i++)
            {
                if (g.arrStyle[i, 0] == s.PadLeft(2, '0'))
                {
                    return g.arrStyle[i, 1];
                }
            }

            return string.Empty;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmZipCode frm = new frmZipCode(txtZip1.Text + txtZip2.Text);
            frm.ShowDialog();
            string fZipCode = frm.rZipCode;     // 郵便番号
            string fZipAdd = frm.rAdd;          // 住所漢字
            string fZipAddFuri = frm.rAddFuri;  // 住所フリガナ
            frm.Dispose();

            if (fZipCode != string.Empty)
            {
                txtZip1.Text = fZipCode.Substring(0, 3);    // 郵便番号
                txtZip2.Text = fZipCode.Substring(3, 4);

                //  ()表記は除去する 2016/06/07
                int zC = fZipAddFuri.IndexOf("(");

                if (zC != -1)
                {
                    fZipAddFuri = fZipAddFuri.Replace(fZipAddFuri.Substring(zC, fZipAddFuri.Length - zC), "");
                }

                // 住所フリガナ：「ｲｶﾆｹｲｻｲｶﾞﾅｲﾊﾞｱｲ」を除去　2016/06/07
                txtAddFuri.Text = Utility.addressUpdate(txtAddFuri.Text, fZipAddFuri.Replace(global.IKAKEISAI_ADD, ""));

                //  ()表記は除去する 2016/06/07
                zC = fZipAdd.IndexOf("（");

                if (zC != -1)
                {
                    fZipAdd = fZipAdd.Replace(fZipAdd.Substring(zC, fZipAdd.Length - zC), "");
                }

                // 住所漢字：「以下に掲載がない場合」を除去　2016/06/07
                string knAdd = (txtAdd.Text.Replace(" ", "").Replace(fZipAdd.Replace(global.IKAKEISAIKN_ADD, ""), "")).Trim();
                txtAdd.Text = fZipAdd.Replace(global.IKAKEISAIKN_ADD, "") + " " + knAdd;
            }
        }

        private void txtMaker_Leave(object sender, EventArgs e)
        {
            TextBox txtbox = (TextBox)sender;

            txtbox.Text = Utility.getStrConv(txtbox.Text);
        }

        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            //カレントデータの更新
            cuDataUpdate(cID[cI]);

            //レコードの移動
            cI = hScrollBar1.Value;
            showOcrData(cI);
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // OCRDataクラス生成
            OCRData ocr = new OCRData();

            // エラーチェックを実行
            if (getErrData(cI, ocr))
            {
                MessageBox.Show("エラーはありませんでした", "エラーチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // データ表示
                showOcrData(cI);
            }
            else
            {
                // カレントインデックスをエラーありインデックスで更新
                cI = ocr._errHeaderIndex;

                // データ表示
                showOcrData(cI);

                // エラー表示
                ErrShow(ocr);
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("表示中のＯＣＲ防犯登録カードデータを削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // データ削除実行
            cuDataDelete(cID[cI]);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // カレントデータ更新
            cuDataUpdate(cID[cI]);

            // データベース更新
            adp.Update(dts.防犯カード);

            // 閉じる
            this.Close();
        }

        private void linkPlus_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        }

        private void btnPlus_Click_1(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < global.ZOOM_MAX)
            {
                leadImg.ScaleFactor += global.ZOOM_STEP;
            }

            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // OCRDataクラス生成
            OCRData ocr = new OCRData();

            // エラーチェックを実行
            if (getErrData(cI, ocr))
            {
                if (MessageBox.Show("エラーはありませんでした。" + Environment.NewLine + "防犯登録カードデータを作成しますか？", "エラーチェック", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }

                // 防犯登録カードデータ作成
                cardDataAdd();

                // 閉じる
                this.Close();
            }
            else
            {
                // カレントインデックスをエラーありインデックスで更新
                cI = ocr._errHeaderIndex;

                // データ表示
                showOcrData(cI);

                // エラー表示
                ErrShow(ocr);
            }
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     防犯登録データ新規登録 </summary>
        ///---------------------------------------------------------------
        private void cardDataAdd()
        {
            this.Cursor = Cursors.WaitCursor;

            cardDataSet cDts = new cardDataSet();
            cardDataSetTableAdapters.防犯登録データTableAdapter cAdp = new cardDataSetTableAdapters.防犯登録データTableAdapter();

            // 2019/06/25 コメント化
            //cAdp.Fill(cDts.防犯登録データ);

            int cnt = 0;

            foreach (var t in dts.防犯カード)
            {
                var c = cDts.防犯登録データ.New防犯登録データRow();
                c.データ区分 = t.データ区分;
                c.画像名 = t.画像名;
                c.登録年 = t.登録年;
                c.登録月 = t.登録月;
                c.登録日 = t.登録日;
                c.登録番号 = t.登録番号;
                c.車体番号 = t.車体番号;
                c.メーカー = t.メーカー;
                c.塗色 = t.塗色;
                c.車種 = t.車種;
                c.郵便番号1 = t.郵便番号1;
                c.郵便番号2 = t.郵便番号2;
                c.車両番号1 = t.車両番号1;
                c.車両番号2 = t.車両番号2;
                c.車名 = t.車名;
                c.住所漢字 = t.住所漢字;
                c.住所1 = t.住所1;
                //c.住所2 = t.住所2;
                c.氏名 = t.氏名;
                c.TEL携帯 = t.TEL携帯;
                c.TEL携帯2 = t.TEL携帯2;
                c.TEL携帯3 = t.TEL携帯3;
                c.PC名 = t.PC名;
                c.CSV作成日 = t.CSV作成日;
                c.備考 = t.備考;
                c.更新年月日 = t.更新年月日;
                c.除外 = global.flgOff;       // 2016/05/30

                cDts.防犯登録データ.Add防犯登録データRow(c);

                cnt++;

                // カードイメージを移動する
                System.IO.File.Move(Properties.Settings.Default.dataPath + t.画像名, Properties.Settings.Default.imgPath + t.画像名);

                // OCRデータを削除する
                t.Delete();
            }

            // データベース更新
            cAdp.Update(cDts.防犯登録データ);
            adp.Update(dts.防犯カード);

            this.Cursor = Cursors.Default;

            // 終了メッセージ
            MessageBox.Show(cnt.ToString("#,##0") + "件の防犯登録データを登録しました", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                txtSharyoNum.Enabled = false;
                txtSharyoNum2.Enabled = false;
                txtCarName.Enabled = false;
            }
            else
            {
                txtSharyoNum.Enabled = true;
                txtSharyoNum2.Enabled = true;
                txtCarName.Enabled = true;
            }
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string zp = txtZip1.Text.Trim() + txtZip2.Text.Trim();

            // 郵便番号から住所を取得
            string zipAdd = string.Empty;
            string zipAddKN = string.Empty;
            int iZ = 0;

            foreach (var t in zipArray)
            {
                string[] zip = t.Split(',');

                if ((txtZip1.Text + txtZip2.Text) == zip[2].Replace("\"", ""))
                {
                    iZ++;

                    if (iZ == 1)
                    {
                        // カナ
                        zipAdd = (zip[4] + " " + zip[5]).Replace("\"", "").Replace(global.IKAKEISAI_ADD, "");

                        //  ()表記は除去する 2016/06/07
                        int zC = zipAdd.IndexOf("(");

                        if (zC != -1)
                        {
                            zipAdd = zipAdd.Replace(zipAdd.Substring(zC, zipAdd.Length - zC), "");
                        }
                        
                        // 漢字
                        zipAddKN = (zip[7] + zip[8]).Replace("\"", "").Replace(global.IKAKEISAIKN_ADD, "");

                        //  ()表記は除去する 2016/06/07
                        zC = zipAddKN.IndexOf("（");

                        if (zC != -1)
                        {
                            zipAddKN = zipAddKN.Replace(zipAddKN.Substring(zC, zipAddKN.Length - zC), "");
                        }
                    }
                    else
                    {
                        // 複数該当した場合
                        break;
                    }
                }
            }

            if (iZ == 1)
            {
                // 単独で該当郵便番号あり
                if (zipAdd != string.Empty)
                {
                    // 住所フリガナ
                    txtAddFuri.Text = Utility.addressUpdate(txtAddFuri.Text, zipAdd);

                    // 住所漢字
                    string knAdd = (txtAdd.Text.Replace(" ", "").Replace(zipAddKN.Replace(" ", ""), "")).Trim();
                    txtAdd.Text = zipAddKN + " " + knAdd;
                }
            }
            else if (iZ > 1)
            {
                // 複数の該当郵便番号あり
                string msg = "郵便番号 " + zp + " は複数の地名が存在します。" + Environment.NewLine;
                msg += "「〒⇔住所」ボタンから該当する地名を選択してください";
                MessageBox.Show(msg,"複数地名あり",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
        }

        private void txtZip1_TextChanged(object sender, EventArgs e)
        {
            string zp = txtZip1.Text.Trim() + txtZip2.Text.Trim();

            if (zp.Length < 7)
            {
                linkLabel5.Enabled = false;
            }
            else
            {
                linkLabel5.Enabled = true;
            }
        }

        private void txtZip2_TextChanged(object sender, EventArgs e)
        {
            string zp = txtZip1.Text.Trim() + txtZip2.Text.Trim();

            if (zp.Length < 7)
            {
                linkLabel5.Enabled = false;
            }
            else
            {
                linkLabel5.Enabled = true;
            }
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        private void linkKengai_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("「" + global.KENGAI_ADD + "」を住所先頭に追加します。よろしいですか","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // ケンガイ追加
            txtAddFuri.Text = kengaiAddress();
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     ケンガイ追加 </summary>
        /// <returns>
        ///     住所フリガナ</returns>
        ///----------------------------------------------------------
        private string kengaiAddress()
        {
            string ad = (global.KENGAI_ADD + " " + txtAddFuri.Text);

            if (ad.Length > 40)
            {
                ad = ad.Substring(0, 40);
            }
            else
            {
                ad = ad.Trim();
            }

            return ad;
        }
    }
}
