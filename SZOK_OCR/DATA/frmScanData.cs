using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SZOK_OCR.Common;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace SZOK_OCR.DATA
{
    public partial class frmScanData : Form
    {
        // 検索ＩＤ
        int dID = 0;

        //string[] zipArray = null;   // 郵便番号配列

        public frmScanData(int sID)
        {
            InitializeComponent();
            
            // 登録済みデータの検索及び編集
            dID = sID;
        }

        Image OcrImg = null;

        // 画像サイズ
        float B_WIDTH = 0.43f;
        float B_HEIGHT = 0.43f;
        float n_width = 0f;
        float n_height = 0f;

        private void frmCorrect_Load(object sender, EventArgs e)
        {
            this.pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);
            
            // キャプション
            this.Text = "自転車防犯登録カード";
                        
            // 指定レコードを表示
            showOcrData(dID);

            // tagを初期化
            this.Tag = string.Empty;
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     伝票画像表示 </summary>
        /// <param name="iX">
        ///     現在の伝票</param>
        /// <param name="tempImgName">
        ///     画像名</param>
        /// ------------------------------------------------------------------------------
        public void ShowImage(string img)
        {
            if (System.IO.File.Exists(img))
            {
                // System.Drawing.Imageを作成する
                OcrImg = Utility.CreateImage(img);
                imgShow(OcrImg, B_WIDTH, B_HEIGHT);
                trackBar1.Enabled = true;
                btnLeft.Enabled = true;
            }
            else
            {
                pictureBox1.Image = null;
                trackBar1.Enabled = false;
                btnLeft.Enabled = false;
            }

        }
        

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
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
            System.Windows.Forms.TextBox txtbox = (System.Windows.Forms.TextBox)sender;

            txtbox.Text = Utility.getStrConv(txtbox.Text);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            // 閉じる
            this.Close();
        }


        private void btnPlus_Click_1(object sender, EventArgs e)
        {
            //if (leadImg.ScaleFactor < global.ZOOM_MAX)
            //{
            //    leadImg.ScaleFactor += global.ZOOM_STEP;
            //}

            //global.miMdlZoomRate = (float)leadImg.ScaleFactor;
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
                //txtSharyoNum.Enabled = false;
                //txtSharyoNum2.Enabled = false;
                //txtCarName.Enabled = false;
            }
            else
            {
                //txtSharyoNum.Enabled = true;
                //txtSharyoNum2.Enabled = true;
                //txtCarName.Enabled = true;
            }
        }

        ///---------------------------------------------------------
        /// <summary>
        ///     画像表示メイン : 2020/04/14 </summary>
        /// <param name="mImg">
        ///     Mat形式イメージ</param>
        /// <param name="w">
        ///     width</param>
        /// <param name="h">
        ///     height</param>
        ///---------------------------------------------------------
        private void imgShow(Image mImg, float w, float h)
        {
            int cWidth = 0;
            int cHeight = 0;

            int pWidth = panel1.Width - 2;
            int pHeight = panel1.Height - 2;

            try
            {
                Bitmap bt = new Bitmap(mImg);

                // Bitmapサイズ
                if (pWidth < (bt.Width * w) || pHeight < (bt.Height * h))
                {
                    cWidth = (int)(bt.Width * w);
                    cHeight = (int)(bt.Height * h);
                }
                else
                {
                    cWidth = pWidth;
                    cHeight = pHeight;
                }

                // Bitmap を生成
                Bitmap canvas = new Bitmap(cWidth, cHeight);

                // ImageオブジェクトのGraphicsオブジェクトを作成する
                Graphics g = Graphics.FromImage(canvas);

                // 画像をcanvasの座標(0, 0)の位置に指定のサイズで描画する
                g.DrawImage(bt, 0, 0, bt.Width * w, bt.Height * h);

                //メモリクリア
                bt.Dispose();
                g.Dispose();

                // PictureBox1に表示する
                pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;
                pictureBox1.Image = canvas;
            }
            catch (Exception ex)
            {
                pictureBox1.Image = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void trackBar1_ValueChanged(object sender, EventArgs e)
        {
            n_width  = B_WIDTH  + (float)trackBar1.Value * 0.05f;
            n_height = B_HEIGHT + (float)trackBar1.Value * 0.05f;

            imgShow(OcrImg, n_width, n_height);
        }

        ///-------------------------------------------------------
        /// <summary>
        ///     画像回転 </summary>
        /// <param name="img">
        ///     Image</param>
        ///-------------------------------------------------------
        private void ImageRotate(Image img)
        {
            Bitmap bmp = (Bitmap)img;

            // 反転せず時計回りに90度回転する
            bmp.RotateFlip(RotateFlipType.Rotate90FlipNone);

            //表示
            pictureBox1.Image = img;
        }

        private void btnLeft_Click(object sender, EventArgs e)
        {
            ImageRotate(pictureBox1.Image);
        }
    }
}
