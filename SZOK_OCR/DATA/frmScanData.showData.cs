using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using SZOK_OCR.Common;


namespace SZOK_OCR.DATA
{
    partial class frmScanData
    {
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     データを画面に表示します </summary>
        /// <param name="iX">
        ///     データインデックス</param>
        ///------------------------------------------------------------------------------------
        private void showOcrData(int iX)
        {
            // SQL Server接続
            var master = new ClsMaster(Properties.Settings.Default.sServerName, Properties.Settings.Default.sLogin,
                                   Properties.Settings.Default.sPass, Properties.Settings.Default.sDatabase);
            
            var r = master.GetData<TblScandata>(iX.ToString());

            // フォーム初期化
            formInitialize();

            // 情報表示
            if (r.DataCategory == global.flgOff)
            {
                radioButton1.Checked = true;
            }
            else if (r.DataCategory == global.flgOn)
            {
                radioButton2.Checked = true;
            }

            txtTourokuNum.Text = r.Number;
            txtShataiNum.Text = r.VehicleIdentificationNumber;
            txtYYMMDD.Text = $"20{r.AddYear}年{r.AddMonth}月{r.AddDay}日";
            txtMaker.Text = r.Maker;
            txtColor.Text = r.Color;

            global g = new global();
            for (int i = 0; i < g.arrStyle.GetLength(0); i++)
            {
                if (g.arrStyle[i, 0] == r.CarModel.ToString().PadLeft(2, '0'))
                {
                    txtStyle.Text = r.CarModel.ToString().PadLeft(2, '0');
                    txtStyleName.Text = g.arrStyle[i, 1];
                    break;
                }
            }

            txtZip1.Text = r.ZipCode1;
            txtZip2.Text = r.ZipCode2;

            txtAddFuri.Text = r.Address1;

            txtFuri.Text = r.Name;
            txtTel.Text = r.Mobile1;
            txtTel2.Text = r.Mobile2;
            txtTel3.Text = r.Mobile3;

            if (r.Memo == null)
            {
                txtMemo.Text = string.Empty;
            }
            else
            {
                txtMemo.Text = r.Memo;
            }

            txtSharyoNum.Text = r.VehicleNumber1;
            txtSharyoNum2.Text = r.VehicleNumber2;
            txtCarName.Text = r.CarName;

            // 画像表示
            ShowImage(Properties.Settings.Default.scanDataPath + r.ImageFileName.ToString());

            linkLabel1.Focus();

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
            txtYYMMDD.BackColor = Color.White;
            txtMaker.BackColor = Color.White;
            txtColor.BackColor = Color.White;
            txtStyle.BackColor = Color.White;
            txtStyleName.BackColor = Color.White;
            txtSharyoNum.BackColor = Color.White;
            txtSharyoNum2.BackColor = Color.White;
            txtCarName.BackColor = Color.White;
            txtZip1.BackColor = Color.White;
            txtZip2.BackColor = Color.White;
            txtAddFuri.BackColor = Color.White;
            txtFuri.BackColor = Color.White;
            txtTel.BackColor = Color.White;
            txtTel2.BackColor = Color.White;
            txtTel3.BackColor = Color.White;
            txtMemo.BackColor = Color.White;

            txtTourokuNum.ForeColor = Color.Navy;
            txtShataiNum.ForeColor = Color.Navy;
            txtYYMMDD.ForeColor = Color.Navy;
            txtMaker.ForeColor = Color.Navy;
            txtColor.ForeColor = Color.Navy;
            txtStyle.ForeColor = Color.Navy;
            txtStyleName.ForeColor = Color.Navy;
            txtSharyoNum.ForeColor = Color.Navy;
            txtSharyoNum2.ForeColor = Color.Navy;
            txtCarName.ForeColor = Color.Navy;
            txtZip1.ForeColor = Color.Navy;
            txtZip2.ForeColor = Color.Navy;
            txtAddFuri.ForeColor = Color.Navy;
            txtFuri.ForeColor = Color.Navy;
            txtTel.ForeColor = Color.Navy;
            txtTel2.ForeColor = Color.Navy;
            txtTel3.ForeColor = Color.Navy;
            txtMemo.ForeColor = Color.Navy;

            lblNoImage.Visible = false;

            dispShowMode();
        }
           

        private void dispShowMode()
        {
            // 情報
            radioButton1.AutoCheck = false;
            radioButton2.AutoCheck = false;
            txtTourokuNum.ReadOnly = true;
            txtShataiNum.ReadOnly = true;
            txtYYMMDD.ReadOnly = true;
            txtStyle.ReadOnly = true;
            txtStyleName.ReadOnly = true;
            txtSharyoNum.ReadOnly = true;
            txtSharyoNum2.ReadOnly = true;
            txtCarName.ReadOnly = true;
            txtZip1.ReadOnly = true;
            txtZip2.ReadOnly = true;
            txtAddFuri.ReadOnly = true;
            txtFuri.ReadOnly = true;
            txtTel.ReadOnly = true;
            txtTel2.ReadOnly = true;
            txtTel3.ReadOnly = true;

            txtMaker.ReadOnly = true;
            txtColor.ReadOnly = true;
            txtMemo.ReadOnly = true;
        }
    }
}
