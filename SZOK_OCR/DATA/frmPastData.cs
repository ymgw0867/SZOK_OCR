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
    public partial class frmPastData : Form
    {
        // 検索ＩＤ
        int dID = 0;

        string EDIT_MODE = "閲覧モードにする";
        string DISP_MODE = "編集モードにする";
        string DELTAG = "delete";
        bool EditMode = false;

        string[] zipArray = null;   // 郵便番号配列

        public frmPastData(int sID)
        {
            InitializeComponent();
            
            // 登録済みデータの検索及び編集
            dID = sID;
        }

        TblRegistrationCard r = null;

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

            // 郵便番号CSV配列読み込み
            Utility.zipCsvLoad(ref zipArray);
            
            // キャプション
            this.Text = "ＯＣＲ防犯登録データ";
                        
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
            n_width = B_WIDTH + (float)trackBar1.Value * 0.05f;
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
            if (System.IO.File.Exists(tempImgName))
            {
                // System.Drawing.Imageを作成する
                OcrImg = Utility.CreateImage(tempImgName);
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
            if (EditMode)
            {
                // カレントデータ更新
                cuDataUpdate();
            }

            // 後片付け
            this.Dispose();
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     防犯カードデータ更新 </summary>
        /// <param name="sID">
        ///     防犯カードデータID</param>
        ///-------------------------------------------------------------------
        private void cuDataUpdate()
        {
            // TblRegistrationCard クラス取得 2026/08/31 追加
            GetTblRegistrationCard();

            // SQL Server更新：2026/08/31
            var master = new ClsMaster(Properties.Settings.Default.sServerName, Properties.Settings.Default.sLogin,
                                   Properties.Settings.Default.sPass, Properties.Settings.Default.sDatabase);

            // 防犯登録データ更新：2026/08/31
            master.UpDate<TblRegistrationCard>(r);
        }

        /// <summary>
        ///    TblRegistrationCard クラス取得 2026/08/31 追加
        /// </summary>
        private void GetTblRegistrationCard()
        {
            // 自転車・原付
            if (radioButton1.Checked)
            {
                r.DataCategory = global.flgOff;
            }
            else if (radioButton2.Checked)
            {
                r.DataCategory = global.flgOn;
            }

            r.AddYear = txtYear.Text;
            r.AddMonth = txtMonth.Text;
            r.AddDay = txtDay.Text;

            r.Number = kanmaDelete(txtTourokuNum.Text);
            r.VehicleIdentificationNumber = kanmaDelete(txtShataiNum.Text);
            r.Maker = kanmaDelete(txtMaker.Text);
            r.Color = kanmaDelete(txtColor.Text);
            r.CarModel = Utility.StrtoInt(txtStyle.Text);
            r.VehicleNumber1 = kanmaDelete(txtSharyoNum.Text);
            r.VehicleNumber2 = kanmaDelete(txtSharyoNum2.Text);
            r.CarName = kanmaDelete(txtCarName.Text);
            r.ZipCode1 = txtZip1.Text;
            r.ZipCode2 = txtZip2.Text;
            r.AddressKanji = kanmaDelete(txtAdd.Text);
            r.Address1 = kanmaDelete(txtAddFuri.Text);
            r.Name = kanmaDelete(txtFuri.Text);
            r.Mobile1 = txtTel.Text;
            r.Mobile2 = txtTel2.Text;
            r.Mobile3 = txtTel3.Text;
            r.Memo = txtMemo.Text;

            if (!checkBox1.Checked)
            {
                r.CsvCreationDate = string.Empty;
            }

            r.UpDate = DateTime.Now;

            // データ除外 2016/05/30
            if (chkJyogai.Checked)
            {
                r.Exception = global.flgOn;
            }
            else
            {
                r.Exception = global.flgOff;
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


        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (linkLabel4.Text == DISP_MODE)
            {
                if (MessageBox.Show("データの編集を可能にしますか", "変更確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }

                if (MessageBox.Show("本当にデータの編集を可能にしますか", "変更確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }

                linkLabel4.Text = EDIT_MODE;
                dispEditMode(); // 編集モードへ
                EditMode = true;   // 編集モードフラグ：2026/08/31
            }
            else if (linkLabel4.Text == EDIT_MODE)
            {
                if (MessageBox.Show("閲覧モードにしますか", "変更確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }

                linkLabel4.Text = DISP_MODE;
                dispShowMode(); // 閲覧モードへ
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
