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
            // OCR防犯テーブル行を取得
            SZOKDataSet.OCR防犯Row r = (SZOKDataSet.OCR防犯Row)dts.OCR防犯.Rows[iX];

            // フォーム初期化
            formInitialize(dID, iX);

            // 情報表示
            txtTourokuNum.Text = r.登録番号;
            txtShataiNum.Text = r.車体番号;
            label5.Text = Properties.Settings.Default.gengou;
            txtYear.Text = r.登録年;
            txtMonth.Text = r.登録月;
            txtDay.Text = r.登録日;
            txtKeisatsuNum.Text = r.警察整理番号;
            txtSeiriNum.Text = r.整理番号;

            if (Utility.StrtoInt(r.メーカー) < cmbMaker.Items.Count)
            {
                cmbMaker.SelectedIndex = Utility.StrtoInt(r.メーカー);
            }
            else
            {
                cmbMaker.SelectedIndex = 0;
            }

            if (Utility.StrtoInt(r.塗色) < cmbColor.Items.Count)
            {
                cmbColor.SelectedIndex = Utility.StrtoInt(r.塗色);
            }
            else
            {
                cmbColor.SelectedIndex = 0;
            }
            
            if (Utility.StrtoInt(r.車種) < cmbStyle.Items.Count)
            {
                cmbStyle.SelectedIndex = Utility.StrtoInt(r.車種);
            }
            else
            {
                cmbStyle.SelectedIndex = 0;
            }

            string zipCode = r.郵便番号.PadLeft(7, ' ');
            txtZip1.Text = zipCode.Substring(0, 3).Trim();
            txtZip2.Text = zipCode.Substring(3, 4).Trim();

            txtFuri.Text = r.フリガナ;
            txtName.Text = r.氏名;
            txtTel.Text = r.TEL携帯1 + r.TEL携帯2 + r.TEL携帯3;

            // エラー情報表示初期化
            lblErrMsg.Visible = false;
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
            txtKeisatsuNum.BackColor = Color.White;
            txtSeiriNum.BackColor = Color.White;
            txtZip1.BackColor = Color.White;
            txtZip2.BackColor = Color.White;
            txtAddFuri.BackColor = Color.White;
            txtAdd.BackColor = Color.White;
            txtFuri.BackColor = Color.White;
            txtName.BackColor = Color.White;
            txtTel.BackColor = Color.White;

            txtTourokuNum.ForeColor = Color.Navy;
            txtShataiNum.ForeColor = Color.Navy;
            txtYear.ForeColor = Color.Navy;
            txtMonth.ForeColor = Color.Navy;
            txtDay.ForeColor = Color.Navy;
            txtKeisatsuNum.ForeColor = Color.Navy;
            txtSeiriNum.ForeColor = Color.Navy;
            txtZip1.ForeColor = Color.Navy;
            txtZip2.ForeColor = Color.Navy;
            txtAddFuri.ForeColor = Color.Navy;
            txtAdd.ForeColor = Color.Navy;
            txtFuri.ForeColor = Color.Navy;
            txtName.ForeColor = Color.Navy;
            txtTel.ForeColor = Color.Navy;

            lblErrMsg.Text = string.Empty;
            lblNoImage.Visible = false;

            // OCR防犯登録データ登録のとき
            if (sID == string.Empty)
            {
                // 情報
                txtTourokuNum.ReadOnly = false;
                txtShataiNum.ReadOnly = false;
                txtYear.ReadOnly = false;
                txtMonth.ReadOnly = false;
                txtDay.ReadOnly = false;
                txtKeisatsuNum.ReadOnly = false;
                txtSeiriNum.ReadOnly = false;
                txtZip1.ReadOnly = false;
                txtZip2.ReadOnly = false;
                txtAddFuri.ReadOnly = false;
                txtAdd.ReadOnly = false;
                txtFuri.ReadOnly = false;
                txtName.ReadOnly = false;
                txtTel.ReadOnly = false;

                // スクロールバー設定
                hScrollBar1.Enabled = true;
                hScrollBar1.Minimum = 0;
                hScrollBar1.Maximum =  dts.OCR防犯.Count - 1;
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
                if ((cIx + 1) == dts.OCR防犯.Count)
                {
                    btnNext.Enabled = false;
                    btnEnd.Enabled = false;
                }

                // その他のボタンを有効とする
                btnErrCheck.Visible = true;
                //btnDataMake.Visible = true;
                btnDel.Visible = true;

                ////エラー情報表示
                //ErrShow();

                //データ数表示
                lblPage.Text = " (" + (cI + 1).ToString() + "/" + dts.OCR防犯.Rows.Count.ToString() + ")";
            }
            else
            {
                // ヘッダ情報
                txtTourokuNum.ReadOnly = false;
                txtShataiNum.ReadOnly = false;
                txtYear.ReadOnly = false;
                txtMonth.ReadOnly = false;
                txtDay.ReadOnly = false;
                txtKeisatsuNum.ReadOnly = false;
                txtSeiriNum.ReadOnly = false;
                txtZip1.ReadOnly = false;
                txtZip2.ReadOnly = false;
                txtAddFuri.ReadOnly = false;
                txtAdd.ReadOnly = false;
                txtFuri.ReadOnly = false;
                txtName.ReadOnly = false;
                txtTel.ReadOnly = false;

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
                btnErrCheck.Visible = false;
                //btnDataMake.Visible = false;
                btnDel.Visible = false;
                
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

                // 対象年月
                if (ocr._errNumber == ocr.eYearMonth)
                {
                    txtYear.BackColor = Color.Yellow;
                    txtMonth.BackColor = Color.Yellow;
                    txtYear.Focus();
                }

                // 対象月
                if (ocr._errNumber == ocr.eMonth)
                {
                    txtMonth.BackColor = Color.Yellow;
                    txtMonth.Focus();
                }

                // 対象日
                if (ocr._errNumber == ocr.eDay)
                {
                    txtDay.BackColor = Color.Yellow;
                    txtDay.Focus();
                }

                // 警察番号
                if (ocr._errNumber == ocr.eKeisatsuNum)
                {
                    txtKeisatsuNum.BackColor = Color.Yellow;
                    txtKeisatsuNum.Focus();
                }

                // 整理番号
                if (ocr._errNumber == ocr.eSeiriNum)
                {
                    txtSeiriNum.BackColor = Color.Yellow;
                    txtSeiriNum.Focus();
                }

                // メーカー
                if (ocr._errNumber == ocr.eMaker)
                {
                    cmbMaker.Focus();
                }

                // 塗色
                if (ocr._errNumber == ocr.eColor)
                {
                    cmbColor.Focus();
                }

                // 車種
                if (ocr._errNumber == ocr.eStyle)
                {
                    cmbStyle.Focus();
                }

                // 郵便番号
                if (ocr._errNumber == ocr.eZip1)
                {
                    txtZip1.BackColor = Color.Yellow;
                    txtZip1.Focus();
                }

                if (ocr._errNumber == ocr.eZip2)
                {
                    txtZip2.BackColor = Color.Yellow;
                    txtZip2.Focus();
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

                // 氏名
                if (ocr._errNumber == ocr.eName)
                {
                    txtName.BackColor = Color.Yellow;
                    txtName.Focus();
                }

            }
        }
    }
}
