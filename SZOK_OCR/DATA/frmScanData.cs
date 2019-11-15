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
    public partial class frmScanData : Form
    {
        cardDataSet dts = new cardDataSet();
        cardDataSetTableAdapters.SCAN_DATATableAdapter adp = new cardDataSetTableAdapters.SCAN_DATATableAdapter();

        // 検索ＩＤ
        int dID = 0;

        string[] zipArray = null;   // 郵便番号配列

        public frmScanData(int sID)
        {
            InitializeComponent();
            
            // 登録済みデータの検索及び編集
            dID = sID;

            // 2019/11/14 対象IDに絞込
            adp.FillByID(dts.SCAN_DATA, dID);
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
            
            // キャプション
            this.Text = "自転車防犯登録カード";
                        
            // 指定レコードを表示
            showOcrData(dID);

            // tagを初期化
            this.Tag = string.Empty;
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            SZOK_OCR.OCR.frmZipCode frm = new SZOK_OCR.OCR.frmZipCode(txtZip1.Text + txtZip2.Text);
            frm.ShowDialog();
            string fZipCode = frm.rZipCode;
            string fZipAdd = frm.rAdd;
            string fZipAddFuri = frm.rAddFuri;
            frm.Dispose();

            if (fZipCode != string.Empty)
            {
                txtZip1.Text = fZipCode.Substring(0, 3);
                txtZip2.Text = fZipCode.Substring(3, 4);
                txtAddFuri.Text = fZipAddFuri;
                txtAdd.Text = fZipAdd;
            }
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
            var r = dts.防犯登録データ.Single(a => a.ID == sID);

            // 自転車・原付
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
            r.住所漢字 = kanmaDelete(txtAdd.Text);
            r.住所1 = kanmaDelete(txtAddFuri.Text);
            r.住所2 = string.Empty;
            r.氏名 = kanmaDelete(txtFuri.Text);
            r.TEL携帯 = txtTel.Text;
            r.TEL携帯2 = txtTel2.Text;
            r.TEL携帯3 = txtTel3.Text;
            r.備考 = txtMemo.Text;
            r.更新年月日 = DateTime.Now;
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
        
        private void txtMaker_Leave(object sender, EventArgs e)
        {
            TextBox txtbox = (TextBox)sender;

            txtbox.Text = Utility.getStrConv(txtbox.Text);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

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

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
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

    }
}
