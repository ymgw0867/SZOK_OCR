using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;


namespace SZOK_OCR.DATA
{
    partial class frmPastData
    {
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     データを画面に表示します </summary>
        /// <param name="iX">
        ///     データインデックス</param>
        ///------------------------------------------------------------------------------------
        private void showOcrData(int iX)
        {
            // 防犯カード行を取得
            var r = dts.防犯登録データ.Single(a => a.ID == iX);

            // フォーム初期化
            formInitialize();

            // 情報表示
            if (r.データ区分 == global.flgOff)
            {
                radioButton1.Checked = true;
            }
            else if (r.データ区分 == global.flgOn)
            {
                radioButton2.Checked = true;
            }

            txtTourokuNum.Text = r.登録番号;
            txtShataiNum.Text = r.車体番号;
            txtYear.Text = r.登録年;
            txtMonth.Text = r.登録月;
            txtDay.Text = r.登録日;
            txtMaker.Text = r.メーカー;
            txtColor.Text = r.塗色;

            global g = new global();
            for (int i = 0; i < g.arrStyle.GetLength(0); i++)
            {
                if (g.arrStyle[i, 0] == r.車種.ToString().PadLeft(2, '0'))
                {
                    txtStyle.Text = r.車種.ToString().PadLeft(2, '0');
                    txtStyleName.Text = g.arrStyle[i, 1];
                    break;
                }
            }

            if (r.Is車両番号1Null())
            {
                txtSharyoNum.Text = "";
            }
            else
            {
                txtSharyoNum.Text = r.車両番号1;
            }

            txtSharyoNum2.Text = r.車両番号2;
            txtCarName.Text = r.車名;

            txtZip1.Text = r.郵便番号1;
            txtZip2.Text = r.郵便番号2;

            txtAddFuri.Text = r.住所1;

            if (r.Is住所漢字Null())
            {
                txtAdd.Text = string.Empty;
            }
            else
            {
                txtAdd.Text = r.住所漢字;
            }

            txtFuri.Text = r.氏名;
            txtTel.Text = r.TEL携帯;
            txtTel2.Text = r.TEL携帯2;
            txtTel3.Text = r.TEL携帯3;

            if (r.Is備考Null())
            {
                txtMemo.Text = string.Empty;
            }
            else
            {
                txtMemo.Text = r.備考;
            }

            // 県警察本部用ＣＳＶデータ作成日
            if (r.CSV作成日 != string.Empty)
            {
                checkBox1.Checked = true;
            }
            else
            {
                checkBox1.Checked = false;
            }

            label22.Text = r.CSV作成日;

            // エラー情報表示初期化
            //lblErrMsg.Visible = false;
            //lblErrMsg.Text = string.Empty;

            // 画像表示
            ShowImage(Properties.Settings.Default.imgPath + r.画像名.ToString());

            linkLabel1.Focus();

            // 除外データのとき
            if (!r.Is除外Null() && r.除外 == global.flgOn)
            {
                // 除外データ
                lblData.Visible = true;
                chkJyogai.Checked = true;
                //chkJyogai.Visible = true;
                //linkLabel2.Visible = false;
            }
            else
            {
                lblData.Visible = false;
                chkJyogai.Checked = false;
                //chkJyogai.Visible = false;
                //linkLabel2.Visible = true;
            }
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     画像を表示する </summary>
        /// <param name="pic">
        ///     pictureBoxオブジェクト</param>
        /// <param name="imgName">
        ///     イメージファイルパス</param>
        /// <param name="fX">
        ///     X方向のスケールファクター</param>
        /// <param name="fY">
        ///     Y方向のスケールファクター</param>
        ///------------------------------------------------------------------------------------
        private void ImageGraphicsPaint(PictureBox pic, string imgName, float fX, float fY, int RectDest, int RectSrc)
        {
            Image _img = Image.FromFile(imgName);
            Graphics g = Graphics.FromImage(pic.Image);

            // 各変換設定値のリセット
            g.ResetTransform();

            // X軸とY軸の拡大率の設定
            g.ScaleTransform(fX, fY);

            // 画像を表示する
            g.DrawImage(_img, RectDest, RectSrc);

            // 現在の倍率,座標を保持する
            global.ZOOM_NOW = fX;
            global.RECTD_NOW = RectDest;
            global.RECTS_NOW = RectSrc;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     フォーム表示初期化 </summary>
        /// <param name="sID">
        ///     過去データ表示時のヘッダID</param>
        /// <param name="cIx">
        ///     勤務票ヘッダカレントレコードインデックス</param>
        ///------------------------------------------------------------------------------------
        private void formInitialize()
        {
            // テキストボックス表示色設定
            txtTourokuNum.BackColor = Color.White;
            txtShataiNum.BackColor = Color.White;
            txtYear.BackColor = Color.White;
            txtMonth.BackColor = Color.White;
            txtDay.BackColor = Color.White;
            txtMaker.BackColor = Color.White;
            txtColor.BackColor = Color.White;
            txtStyle.BackColor = Color.White;
            txtStyleName.BackColor = Color.White;
            txtSharyoNum.BackColor = Color.White;
            txtCarName.BackColor = Color.White;
            txtZip1.BackColor = Color.White;
            txtZip2.BackColor = Color.White;
            txtAddFuri.BackColor = Color.White;
            txtAdd.BackColor = Color.White;
            txtFuri.BackColor = Color.White;
            txtTel.BackColor = Color.White;
            txtTel2.BackColor = Color.White;
            txtTel3.BackColor = Color.White;
            txtMemo.BackColor = Color.White;

            txtTourokuNum.ForeColor = Color.Navy;
            txtShataiNum.ForeColor = Color.Navy;
            txtYear.ForeColor = Color.Navy;
            txtMonth.ForeColor = Color.Navy;
            txtDay.ForeColor = Color.Navy;
            txtMaker.ForeColor = Color.Navy;
            txtColor.ForeColor = Color.Navy;
            txtStyle.ForeColor = Color.Navy;
            txtStyleName.ForeColor = Color.Navy;
            txtSharyoNum.ForeColor = Color.Navy;
            txtCarName.ForeColor = Color.Navy;
            txtZip1.ForeColor = Color.Navy;
            txtZip2.ForeColor = Color.Navy;
            txtAddFuri.ForeColor = Color.Navy;
            txtAdd.ForeColor = Color.Navy;
            txtFuri.ForeColor = Color.Navy;
            txtTel.ForeColor = Color.Navy;
            txtTel2.ForeColor = Color.Navy;
            txtTel3.ForeColor = Color.Navy;
            txtMemo.ForeColor = Color.Navy;

            //lblErrMsg.Text = string.Empty;
            lblNoImage.Visible = false;

            dispShowMode();

            linkLabel4.Text = DISP_MODE;
        }

        private void dispEditMode()
        {
            lblMode.Text = "現在、編集モードです";
            lblMode.ForeColor = Color.Red;

            // 情報
            radioButton1.AutoCheck = true;
            radioButton2.AutoCheck = true;
            txtTourokuNum.ReadOnly = false;
            txtShataiNum.ReadOnly = false;
            txtYear.ReadOnly = false;
            txtMonth.ReadOnly = false;
            txtDay.ReadOnly = false;
            txtStyle.ReadOnly = false;
            txtStyleName.ReadOnly = false;
            txtSharyoNum.ReadOnly = false;
            txtSharyoNum2.ReadOnly = false;
            txtCarName.ReadOnly = false;
            txtZip1.ReadOnly = false;
            txtZip2.ReadOnly = false;
            txtAddFuri.ReadOnly = false;
            txtAdd.ReadOnly = false;
            txtFuri.ReadOnly = false;
            txtTel.ReadOnly = false;
            txtTel2.ReadOnly = false;
            txtTel3.ReadOnly = false;

            txtMaker.ReadOnly = false;
            txtColor.ReadOnly = false;
            txtMemo.ReadOnly = false;
            checkBox1.AutoCheck = true;

            button1.Visible = true;
            //linkLabel2.Visible = true;

            chkJyogai.AutoCheck = true;
        }

        private void dispShowMode()
        {
            lblMode.Text = "現在、閲覧モードです";
            lblMode.ForeColor = Color.SteelBlue;

            // 情報
            radioButton1.AutoCheck = false;
            radioButton2.AutoCheck = false;
            txtTourokuNum.ReadOnly = true;
            txtShataiNum.ReadOnly = true;
            txtYear.ReadOnly = true;
            txtMonth.ReadOnly = true;
            txtDay.ReadOnly = true;
            txtStyle.ReadOnly = true;
            txtStyleName.ReadOnly = true;
            txtSharyoNum.ReadOnly = true;
            txtSharyoNum2.ReadOnly = true;
            txtCarName.ReadOnly = true;
            txtZip1.ReadOnly = true;
            txtZip2.ReadOnly = true;
            txtAddFuri.ReadOnly = true;
            txtAdd.ReadOnly = true;
            txtFuri.ReadOnly = true;
            txtTel.ReadOnly = true;
            txtTel2.ReadOnly = true;
            txtTel3.ReadOnly = true;

            txtMaker.ReadOnly = true;
            txtColor.ReadOnly = true;
            txtMemo.ReadOnly = true;

            checkBox1.AutoCheck = false;

            button1.Visible = false;
            linkLabel2.Visible = false;

            chkJyogai.AutoCheck = false;
        }
    }
}
