using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using SZOK_OCR.Common;

namespace SZOK_OCR.OCR
{
    partial class frmCorrect
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
            szokDataSet.防犯カードRow r = dts.防犯カード.Single(a => a.ID == cID[iX]);

            // フォーム初期化
            formInitialize(dID, iX);

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
            
            txtStyle.Text = r.車種.ToString().PadLeft(2, '0');
            txtStyleName.Text = string.Empty;

            global g = new global();
            for (int i = 0; i < g.arrStyle.GetLength(0); i++)
            {
                if (g.arrStyle[i, 0] == r.車種.ToString().PadLeft(2, '0'))
                {
                    txtStyleName.Text = g.arrStyle[i, 1];
                    break;
                }
            }
            
            txtSharyoNum.Text = r.車両番号1;
            txtSharyoNum2.Text = r.車両番号2;
            txtCarName.Text = r.車名;

            txtZip1.Text = r.郵便番号1;
            txtZip2.Text = r.郵便番号2;

            txtAddFuri.Text = r.住所1;

            string zipAdd = string.Empty;
            string zipAddKN = string.Empty;

            // 郵便番号から住所を取得
            foreach (var t in zipArray)
            {
                string[] zip = t.Split(',');

                if ((txtZip1.Text + txtZip2.Text) == zip[2].Replace("\"", ""))
                {
                    zipAdd = (zip[4] + " " + zip[5]).Replace("\"", "");
                    zipAddKN = (zip[7] + " " + zip[8]).Replace("\"", "");
                    break;
                }
            }

            if (r.Is住所漢字Null())
            {
                txtAdd.Text = string.Empty;
            }
            else
            {
                if (r.住所漢字 == string.Empty)
                {
                    txtAdd.Text = zipAddKN;
                }
                else
                {
                    txtAdd.Text = r.住所漢字;
                }
            }

            //// 住所と郵便番号が一致しているか？
            //string str1 = Utility.strSmallTolarge(Utility.getStrConv(txtAddFuri.Text.Replace(" ", "").Trim()));
            //string str2 = Utility.strSmallTolarge((Utility.getStrConv(zipAdd.Replace(" ", "").Trim())));

            //if (!str1.Contains(str2))
            //{
            //    // 一致していないときバックカラーをピンクにする
            //    txtAddFuri.BackColor = Color.LightPink;
            //}
            //else
            //{
            //    // 一致しているときバックカラーを白にする
            //    txtAddFuri.BackColor = Color.White;
            //}

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

            // 画面確認チェック 2016.01/25
            if (r.Is確認Null())
            {
                checkBox1.Checked = false;
                label22.Visible = false;
            }
            else if (r.確認 == global.flgOff)
            {
                checkBox1.Checked = false;
                label22.Visible = false;
            }
            else
            {
                checkBox1.Checked = true;
                label22.Visible = true;
            }

            // エラー情報表示初期化
            //lblErrMsg.Visible = false;
            lblErrMsg.Text = string.Empty;

            // 画像表示
            ShowImage(global.pblImagePath + r.画像名.ToString());
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
        private void formInitialize(string sID, int cIx)
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
            checkBox1.BackColor = SystemColors.Control;

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

            lblErrMsg.Text = string.Empty;
            lblNoImage.Visible = false;

            // 情報
            txtTourokuNum.ReadOnly = false;
            txtShataiNum.ReadOnly = false;
            txtYear.ReadOnly = false;
            txtMonth.ReadOnly = false;
            txtDay.ReadOnly = false;
            txtStyle.ReadOnly = false;
            txtStyleName.ReadOnly = false;
            txtSharyoNum.ReadOnly = false;
            txtCarName.ReadOnly = false;
            txtZip1.ReadOnly = false;
            txtZip2.ReadOnly = false;
            txtAddFuri.ReadOnly = false;
            txtAdd.ReadOnly = false;
            txtFuri.ReadOnly = false;
            txtTel.ReadOnly = false;
            txtTel2.ReadOnly = false;
            txtTel3.ReadOnly = false;

            // 2016/05/30
            linkKengai.Visible = false;

            // OCR防犯登録データ登録のとき
            if (sID == string.Empty)
            {
                // スクロールバー設定
                hScrollBar1.Enabled = true;
                hScrollBar1.Minimum = 0;
                hScrollBar1.Maximum =  dts.防犯カード.Count - 1;
                hScrollBar1.Value = cIx;
                hScrollBar1.LargeChange = 1;
                hScrollBar1.SmallChange = 1;

                //移動ボタン制御
                btnFirst.Enabled = true;
                btnNext.Enabled = true;
                btnBefore.Enabled = true;
                btnEnd.Enabled = true;

                //最初のレコード
                if (cIx == 0)
                {
                    btnBefore.Enabled = false;
                    btnFirst.Enabled = false;
                }

                //最終レコード
                if ((cIx + 1) == cID.Length)
                {
                    btnNext.Enabled = false;
                    btnEnd.Enabled = false;
                }

                // その他のボタンを有効とする
                linkLabel3.Visible = true;
                linkLabel2.Visible = true;

                ////エラー情報表示
                //ErrShow();

                //データ数表示
                lblPage.Text = " (" + (cI + 1).ToString() + "/" + cID.Length.ToString() + ")";
            }
            else
            {
                // スクロールバー設定
                hScrollBar1.Enabled = true;
                hScrollBar1.Minimum = 0;
                hScrollBar1.Maximum = 0;
                hScrollBar1.Value = 0;
                hScrollBar1.LargeChange = 1;
                hScrollBar1.SmallChange = 1;

                //移動ボタン制御
                btnFirst.Enabled = false;
                btnNext.Enabled = false;
                btnBefore.Enabled = false;
                btnEnd.Enabled = false;

                // その他のボタンを無効とする
                linkLabel3.Visible = false;
                linkLabel2.Visible = false;
                
                //データ数表示
                lblPage.Text = string.Empty;
            }
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     エラー表示 </summary>
        /// <param name="ocr">
        ///     OCRDATAクラス</param>
        ///------------------------------------------------------------------------------------
        private void ErrShow(OCRData ocr)
        {
            if (ocr._errNumber != ocr.eNothing)
            {
                lblErrMsg.Visible = true;
                lblErrMsg.Text = ocr._errMsg;

                // 登録番号
                if (ocr._errNumber == ocr.eTourokuNum)
                {
                    txtTourokuNum.BackColor = Color.Yellow;
                    txtTourokuNum.Focus();
                }

                // 車体番号
                if (ocr._errNumber == ocr.eShataiNum)
                {
                    txtShataiNum.BackColor = Color.Yellow;
                    txtShataiNum.Focus();
                }

                // 登録年
                if (ocr._errNumber == ocr.eYear)
                {
                    txtYear.BackColor = Color.Yellow;
                    txtYear.Focus();
                }

                // 登録月
                if (ocr._errNumber == ocr.eMonth)
                {
                    txtMonth.BackColor = Color.Yellow;
                    txtMonth.Focus();
                }

                // 登録日
                if (ocr._errNumber == ocr.eDay)
                {
                    txtDay.BackColor = Color.Yellow;
                    txtDay.Focus();
                }

                // メーカー
                if (ocr._errNumber == ocr.eMaker)
                {
                    txtMaker.BackColor = Color.Yellow;
                    txtMaker.Focus();
                }

                // 塗色
                if (ocr._errNumber == ocr.eColor)
                {
                    txtColor.BackColor = Color.Yellow;
                    txtColor.Focus();
                }

                // 車種
                if (ocr._errNumber == ocr.eStyle)
                {
                    txtStyle.BackColor = Color.Yellow;
                    txtStyle.Focus();
                }
                
                // 車名
                if (ocr._errNumber == ocr.eCarName)
                {
                    txtCarName.BackColor = Color.Yellow;
                    txtCarName.Focus();
                }

                // 郵便番号
                if (ocr._errNumber == ocr.eZip1)
                {
                    txtZip1.BackColor = Color.Yellow;
                    txtZip2.BackColor = Color.Yellow;
                    txtZip1.Focus();
                    linkKengai.Visible = true;
                }

                // 住所フリガナ
                if (ocr._errNumber == ocr.eAddFuri)
                {
                    txtAddFuri.BackColor = Color.Yellow;
                    txtAddFuri.Focus();
                }

                // 住所
                if (ocr._errNumber == ocr.eAdd)
                {
                    txtAdd.BackColor = Color.Yellow;
                    txtAdd.Focus();
                }

                // フリガナ
                if (ocr._errNumber == ocr.eFuri)
                {
                    txtFuri.BackColor = Color.Yellow;
                    txtFuri.Focus();
                }

                // TEL/携帯番号
                if (ocr._errNumber == ocr.eTel)
                {
                    txtTel.BackColor = Color.Yellow;
                    txtTel2.BackColor = Color.Yellow;
                    txtTel3.BackColor = Color.Yellow;
                    txtTel.Focus();
                }

                // 確認チェック
                if (ocr._errNumber == ocr.eCheck)
                {
                    checkBox1.BackColor = Color.Yellow;
                }
            }
        }
    }
}
