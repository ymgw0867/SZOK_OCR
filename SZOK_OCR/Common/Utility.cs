﻿using System;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.Text.RegularExpressions;


namespace SZOK_OCR.Common
{
    class Utility
    {
        /// ---------------------------------------------------------------------
        /// <summary>
        ///     ウィンドウ最小サイズの設定 </summary>
        /// <param name="tempFrm">
        ///     対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">
        ///     width</param>
        /// <param name="hSize">
        ///     Height</param>
        /// ---------------------------------------------------------------------
        public static void WindowsMinSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MinimumSize = new Size(wSize, hSize);
        }

        /// ---------------------------------------------------------------------
        /// <summary>
        ///     ウィンドウ最小サイズの設定 </summary>
        /// <param name="tempFrm">
        ///     対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">
        ///     width</param>
        /// <param name="hSize">
        ///     height</param>
        /// --------------------------------------------------------------------
        public static void WindowsMaxSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MaximumSize = new Size(wSize, hSize);
        }

        /// <summary>
        /// フォームのデータ登録モード
        /// </summary>
        public class frmMode
        {
            public int Mode { get; set; }
            public string ID { get; set; }
            public int rowIndex { get; set; }
            public int closeMode { get; set; }
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     休日コンボボックスクラス </summary>
        /// -------------------------------------------------------------------
        public class comboHoliday
        {
            public string Date { get; set; }
            public string Name { get; set; }

            /// -------------------------------------------------------------------------
            /// <summary>
            ///     休日コンボボックスデータロード </summary>
            /// <param name="tempBox">
            ///     ロード先コンボボックスオブジェクト名</param>
            /// -------------------------------------------------------------------------
            public static void Load(ComboBox tempBox)
            {

                // 休日配列
                string[] sDay = {"01/01元旦", "     成人の日", "02/11建国記念の日", "     春分の日", "04/29昭和の日",
                            "05/03憲法記念日","05/04みどりの日","05/05こどもの日","     海の日","     敬老の日",
                            "     秋分の日","     体育の日","11/03文化の日","11/23勤労感謝の日","12/23天皇誕生日",
                            "     振替休日","     国民の休日","     土曜日","     年末年始休暇","     夏季休暇"}; 

                try
                {
                    comboHoliday cmb1;

                    tempBox.Items.Clear();
                    tempBox.DisplayMember = "Name";
                    tempBox.ValueMember = "Date";

                    foreach (var a in sDay)
                    {
                        cmb1 = new comboHoliday();
                        cmb1.Date = a.Substring(0, 5);
                        int s = a.Length;
                        cmb1.Name = a.Substring(5, s - 5);
                        tempBox.Items.Add(cmb1);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "休日コンボボックスロード");
                }
            }

            /// ---------------------------------------------------------------------
            /// <summary>
            ///     休日コンボ表示 </summary>
            /// <param name="tempBox">
            ///     コンボボックスオブジェクト</param>
            /// <param name="dt">
            ///     月日</param>
            /// ---------------------------------------------------------------------
            public static void selectedIndex(ComboBox tempBox, string dt)
            {
                comboHoliday cmbS = new comboHoliday();
                Boolean Sh = false;

                for (int iX = 0; iX <= tempBox.Items.Count - 1; iX++)
                {
                    tempBox.SelectedIndex = iX;
                    cmbS = (comboHoliday)tempBox.SelectedItem;

                    if (cmbS.Date == dt)
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempBox.SelectedIndex = -1;
                }
            }
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     文字列の値が数字かチェックする </summary>
        /// <param name="tempStr">
        ///     検証する文字列</param>
        /// <returns>
        ///     数字:true,数字でない:false</returns>
        /// ------------------------------------------------------------------------------
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }
        
        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     emptyを"0"に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        /// ------------------------------------------------------------------------------
        public static string EmptytoZero(string tempStr)
        {
            if (tempStr == string.Empty)
            {
                return "0";
            }
            else
            {
                return tempStr;
            }
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのとき文字型値を返す</returns>
        /// ------------------------------------------------------------------------------
        public static string NulltoStr(string tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                return tempStr;
            }
        }

        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        /// -------------------------------------------------------------------------------
        public static string NulltoStr(object tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                if (tempStr == DBNull.Value)
                {
                    return string.Empty;
                }
                else
                {
                    return (string)tempStr.ToString();
                }
            }
        }

        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        /// -------------------------------------------------------------------------------
        public static string nulltoStr2(object tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                if (tempStr == DBNull.Value)
                {
                    return string.Empty;
                }
                else
                {
                    return (string)tempStr.ToString();
                }
            }
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     文字型をIntへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     Int型の値</returns>
        /// --------------------------------------------------------------------------------
        public static int StrtoInt(string tempStr)
        {
            if (NumericCheck(tempStr)) return int.Parse(tempStr);
            else return 0;
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     文字型をDoubleへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     double型の値</returns>
        /// --------------------------------------------------------------------------------
        public static double StrtoDouble(string tempStr)
        {
            if (NumericCheck(tempStr)) return double.Parse(tempStr);
            else return 0;
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     経過時間を返す </summary>
        /// <param name="s">
        ///     開始時間</param>
        /// <param name="e">
        ///     終了時間</param>
        /// <returns>
        ///     経過時間</returns>
        /// --------------------------------------------------------------------------------
        public static TimeSpan GetTimeSpan(DateTime s, DateTime e)
        {
            TimeSpan ts;
            if (s > e)
            {
                TimeSpan j = new TimeSpan(24, 0, 0);
                ts = e + j - s;
            }
            else
            {
                ts = e - s;
            }

            return ts;
        }

        /// ------------------------------------------------------------------------
        /// <summary>
        ///     指定した精度の数値に切り捨てます。</summary>
        /// <param name="dValue">
        ///     丸め対象の倍精度浮動小数点数。</param>
        /// <param name="iDigits">
        ///     戻り値の有効桁数の精度。</param>
        /// <returns>
        ///     iDigits に等しい精度の数値に切り捨てられた数値。</returns>
        /// ------------------------------------------------------------------------
        public static double ToRoundDown(double dValue, int iDigits)
        {
            double dCoef = System.Math.Pow(10, iDigits);

            return dValue > 0 ? System.Math.Floor(dValue * dCoef) / dCoef :
                                System.Math.Ceiling(dValue * dCoef) / dCoef;
        }


        //部門コンボボックスクラス
        public class ComboBumon
        {
            public string ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string NameShow { get; set; }
            public string code { get; set; }
            
            ////部門マスターロード
            //public static void load(ComboBox tempObj, int tempLen, string dbName)
            //{
            //    try
            //    {
            //        ComboBumon cmb1;
            //        string sqlSTRING = string.Empty;
            //        dbControl.DataControl sdcon = new dbControl.DataControl(dbName);
            //        OleDbDataReader dR;

            //        sqlSTRING += "select DepartmentID,DepartmentCode,DepartmentName from tbDepartment ";
            //        sqlSTRING += "where DepartmentID <> 1 ";
            //        sqlSTRING += "order by DepartmentCode ";

            //        //scom.CommandText = sqlSTRING;

            //        //データリーダーを取得する
            //        //dR = scom.ExecuteReader();
            //        dR = sdcon.FreeReader(sqlSTRING);

            //        tempObj.Items.Clear();
            //        tempObj.DisplayMember = "Name";
            //        tempObj.ValueMember = "ID";

            //        while (dR.Read())
            //        {
            //            cmb1 = new ComboBumon();
            //            cmb1.ID = dR["DepartmentCode"].ToString();
            //            //cmb1.Name = string.Format("{0:000000000000000}", Int64.Parse(dR["DepartmentCode"].ToString())).Substring(15 - tempLen, tempLen) + " " + dR["DepartmentName"].ToString() + "";

            //            if (Utility.NumericCheck(dR["DepartmentCode"].ToString()))
            //                cmb1.Name = int.Parse(dR["DepartmentCode"].ToString()).ToString().PadLeft(tempLen, '0') + " " + dR["DepartmentName"].ToString() + "";
            //            else cmb1.Name = (dR["DepartmentCode"].ToString() + "    ").Substring(0, tempLen) + " " + dR["DepartmentName"].ToString() + "";

            //            cmb1.NameShow = dR["DepartmentName"].ToString() + "";
            //            tempObj.Items.Add(cmb1);
            //        }

            //        dR.Close();
            //        sdcon.Close();
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message, "部門コンボボックスロード");
            //    }

            //}

            //部門コンボ表示
            public static void selectedIndex(ComboBox tempObj, int id)
            {
                ComboBumon cmbS = new ComboBumon();
                Boolean Sh;

                Sh = false;

                for (int iX = 0; iX <= tempObj.Items.Count - 1; iX++)
                {
                    tempObj.SelectedIndex = iX;
                    cmbS = (ComboBumon)tempObj.SelectedItem;

                    if (cmbS.ID == id.ToString())
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempObj.SelectedIndex = -1;
                }

            }
        }

        // 社員コンボボックスクラス
        public class ComboShain
        {
            public int ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }
            public int YakushokuType { get; set; }
            public string BumonName { get; set; }
            public string BumonCode { get; set; }

            //// 社員マスターロード
            //public static void load(ComboBox tempObj, string dbName)
            //{
            //    try
            //    {
            //        ComboShain cmb1;
            //        string sqlSTRING = string.Empty;
            //        dbControl.DataControl dCon = new dbControl.DataControl(dbName);
            //        OleDbDataReader dR;

            //        sqlSTRING += "select Id,Code, Sei, Mei, YakushokuType from Shain ";
            //        sqlSTRING += "where Shurojokyo = 1 ";
            //        sqlSTRING += "order by Code";

            //        //データリーダーを取得する
            //        dR = dCon.FreeReader(sqlSTRING);

            //        tempObj.Items.Clear();
            //        tempObj.DisplayMember = "DisplayName";
            //        tempObj.ValueMember = "code";

            //        while (dR.Read())
            //        {
            //            cmb1 = new ComboShain();
            //            cmb1.ID = int.Parse(dR["Id"].ToString());
            //            cmb1.DisplayName = dR["Code"].ToString().Trim() + " " + dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
            //            cmb1.Name = dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
            //            cmb1.code = (dR["Code"].ToString() + "").Trim();
            //            cmb1.YakushokuType = int.Parse(dR["YakushokuType"].ToString());
            //            tempObj.Items.Add(cmb1);
            //        }

            //        dR.Close();
            //        dCon.Close();
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message, "社員コンボボックスロード");
            //    }

            //}

            // パートタイマーロード
            //public static void loadPart(ComboBox tempObj, string dbName)
            //{
            //    try
            //    {
            //        ComboShain cmb1;
            //        string sqlSTRING = string.Empty;
            //        dbControl.DataControl dCon = new dbControl.DataControl(dbName);
            //        OleDbDataReader dR;
            //        sqlSTRING += "select Bumon.Code as bumoncode,Bumon.Name as bumonname,Shain.Id as shainid,";
            //        sqlSTRING += "Shain.Code as shaincode,Shain.Sei,Shain.Mei, Shain.YakushokuType ";
            //        sqlSTRING += "from Shain left join Bumon ";
            //        sqlSTRING += "on Shain.BumonId = Bumon.Id ";
            //        sqlSTRING += "where Shurojokyo = 1 and YakushokuType = 1 ";
            //        sqlSTRING += "order by Shain.Code";
                    
            //        //sqlSTRING += "select Id,Code, Sei, Mei, YakushokuType from Shain ";
            //        //sqlSTRING += "where Shurojokyo = 1 and YakushokuType = 1 ";
            //        //sqlSTRING += "order by Code";

            //        //データリーダーを取得する
            //        dR = dCon.FreeReader(sqlSTRING);

            //        tempObj.Items.Clear();
            //        tempObj.DisplayMember = "DisplayName";
            //        tempObj.ValueMember = "code";

            //        while (dR.Read())
            //        {
            //            cmb1 = new ComboShain();
            //            cmb1.ID = int.Parse(dR["shainid"].ToString());
            //            cmb1.DisplayName = dR["shaincode"].ToString().Trim() + " " + dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
            //            cmb1.Name = dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
            //            cmb1.code = (dR["shaincode"].ToString() + "").Trim();
            //            cmb1.YakushokuType = int.Parse(dR["YakushokuType"].ToString());
            //            cmb1.BumonCode = dR["bumoncode"].ToString().PadLeft(3, '0');
            //            cmb1.BumonName = dR["bumonname"].ToString();
            //            tempObj.Items.Add(cmb1);
            //        }

            //        dR.Close();
            //        dCon.Close();
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message, "社員コンボボックスロード");
            //    }

            //}
        }


        // データ領域コンボボックスクラス
        public class ComboDataArea
        {
            public string ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }

            //// データ領域ロード
            //public static void load(ComboBox tempObj)
            //{
            //    dbControl.DataControl dcon = new dbControl.DataControl(Properties.Settings.Default.SQLDataBase);
            //    OleDbDataReader dR = null;

            //    try
            //    {
            //        ComboDataArea cmb;

            //        // データリーダー取得
            //        string mySql = string.Empty;
            //        mySql += "SELECT * FROM Common_Unit_DataAreaInfo ";
            //        mySql += "where CompanyTerm = " + DateTime.Today.Year.ToString();
            //        dR = dcon.FreeReader(mySql);

            //        //会社情報がないとき
            //        if (!dR.HasRows)
            //        {
            //            MessageBox.Show("会社領域情報が存在しません", "会社領域選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //            return;
            //        }

            //        // コンボボックスにアイテムを追加します
            //        tempObj.Items.Clear();
            //        tempObj.DisplayMember = "DisplayName";

            //        while (dR.Read())
            //        {
            //            cmb = new ComboDataArea();
            //            // "CompanyCode"が数字のレコードを対象とする
            //            if (Utility.NumericCheck(dR["CompanyCode"].ToString()))
            //            {
            //                cmb.DisplayName = dR["CompanyName"].ToString().Trim();
            //                cmb.ID = dR["Name"].ToString().Trim();
            //                cmb.code = dR["CompanyCode"].ToString().Trim();
            //                tempObj.Items.Add(cmb);
            //            }
            //        }
            //    }
            //    catch (Exception e)
            //    {
            //        MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            //    }
            //    finally
            //    {
            //        if (!dR.IsClosed) dR.Close();
            //        dcon.Close();
            //    }

            //}

        }


        ///--------------------------------------------------------
        /// <summary>
        /// 会社情報より部門コード桁数、社員コード桁数を取得
        /// </summary>
        /// -------------------------------------------------------
        public class BumonShainKetasu
        {
            public string ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }

            //// 会社情報取得
            //public static void GetKetasu(string dbName)
            //{
            //    dbControl.DataControl dcon = new dbControl.DataControl(dbName);
            //    OleDbDataReader dR = null;

            //    try
            //    {
            //        // データリーダー取得
            //        string mySql = string.Empty;
            //        mySql += "SELECT BumonCodeKeta,ShainCodeKeta FROM Kaisha ";
            //        dR = dcon.FreeReader(mySql);

            //        // 部門コード桁数、社員コード桁数を取得
            //        while (dR.Read())
            //        {
            //            global.ShozokuLength = int.Parse(dR["BumonCodeKeta"].ToString());
            //            global.ShainLength = int.Parse(dR["ShainCodeKeta"].ToString());
            //        }
            //    }
            //    catch (Exception e)
            //    {
            //        MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            //    }
            //    finally
            //    {
            //        if (!dR.IsClosed) dR.Close();
            //        dcon.Close();
            //    }

            //}
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     任意のディレクトリのファイルを削除する </summary>
        /// <param name="sPath">
        ///     指定するディレクトリ</param>
        /// <param name="sFileType">
        ///     ファイル名及び形式</param>
        /// --------------------------------------------------------------------
        public static void FileDelete(string sPath, string sFileType)
        {
            //sFileTypeワイルドカード"*"は、すべてのファイルを意味する
            foreach (string files in System.IO.Directory.GetFiles(sPath, sFileType))
            {
                // ファイルを削除する
                System.IO.File.Delete(files);
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     文字列を指定文字数をＭＡＸとして返します</summary>
        /// <param name="s">
        ///     文字列</param>
        /// <param name="n">
        ///     文字数</param>
        /// <returns>
        ///     文字数範囲内の文字列</returns>
        /// --------------------------------------------------------------------
        public static string GetStringSubMax(string s, int n)
        {
            string val = string.Empty;
            if (s.Length > n) val = s.Substring(0, n);
            else val = s;

            return val;
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     8ケタ左詰め右空白埋めの給与大臣検索用の社員コード文字列を返す
        /// </summary>
        /// <param name="sCode">
        ///     コード</param>
        /// <returns>
        ///     給与大臣検索用の社員コード文字列</returns>
        /// --------------------------------------------------------------------
        public static string bldShainCode(string sCode)
        {
            return sCode.PadLeft(4, '0').PadRight(8, ' ').Substring(0, 8);
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     郵便番号CSVを配列に読み込む </summary>
        /// <param name="z">
        ///     郵便番号CSV配列</param>
        ///--------------------------------------------------------------
        public static void zipCsvLoad(ref string[] z)
        {
            szokDataSetTableAdapters.環境設定TableAdapter adp = new szokDataSetTableAdapters.環境設定TableAdapter();
            szokDataSet dts = new szokDataSet();
            adp.Fill(dts.環境設定);

            var s = dts.環境設定.Single(a => a.ID == global.configKEY);

            if (!System.IO.File.Exists(s.郵便番号データパス))
            {
                return;
            }

            // 郵便番号CSV読み込み
            z = System.IO.File.ReadAllLines(s.郵便番号データパス, Encoding.Default);
        }


        ///---------------------------------------------------------------------------
        /// <summary>
        ///     国内電話番号および携帯番号の正規かチェック </summary>
        /// <param name="tVal">
        ///     対象となるTEL/携帯番号</param>
        /// <returns>true:正しい、false:正しくない</returns>
        ///---------------------------------------------------------------------------
        public static bool regexTelNum(string tVal)
        {
            if (!Regex.IsMatch(tVal, @"^\d{2}-\d{4}-\d{4}$"))
            {
                if (!Regex.IsMatch(tVal, @"^\d{3}-\d{3}-\d{4}$"))
                {
                    if (!Regex.IsMatch(tVal, @"^\d{4}-\d{2}-\d{4}$"))
                    {
                        if (!Regex.IsMatch(tVal, @"^\d{5}-\d{1}-\d{4}$"))
                        {
                            // 携帯番号
                            if (!Regex.IsMatch(tVal, @"^\d{3}-\d{4}-\d{4}$"))
                            {
                                return false;
                            }

                        }
                    }
                }
            }

            return true;
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     全角を半角に変換、ひらがなをカタカナに変換する </summary>
        /// <param name="s">
        ///     変換する文字列</param>
        /// <returns>
        ///     変換結果</returns>
        ///---------------------------------------------------------------------
        public static string getStrConv(string s)
        {
            //全角を半角に変換する
            string s4 = Microsoft.VisualBasic.Strings.StrConv(s, Microsoft.VisualBasic.VbStrConv.Narrow, 0x411);

            //ひらがなをカタカナに変換し、全角を半角に変換する
            s4 = Microsoft.VisualBasic.Strings.StrConv(s4, Microsoft.VisualBasic.VbStrConv.Katakana |
                     Microsoft.VisualBasic.VbStrConv.Narrow, 0x411);

            return s4;
        }


        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVファイルを出力する</summary>
        /// <param name="sPath">
        ///     出力するパス</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        /// <param name="sFileName">
        ///     CSVファイル名</param>
        ///----------------------------------------------------------------------------
        public static void csvFileWrite(string sPath, string[] arrayData, string sFileName)
        {
            // ファイル名
            string outFileName = sPath + sFileName + ".csv";

            // 出力ファイルが存在するとき
            if (System.IO.File.Exists(outFileName))
            {
                // リネーム付加文字列（タイムスタンプ）
                string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
                                     DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
                                     DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

                // リネーム後ファイル名
                string reFileName = sPath + sFileName + newFileName + ".csv";

                // 既存のファイルをリネーム
                System.IO.File.Move(outFileName, reFileName);
            }

            // CSVファイル出力：文字エンコーディングをshift-jisに変更 2016/06/17
            //System.IO.File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding("utf-8"));
            System.IO.File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding("shift-jis"));
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     半角カナの小文字を大文字に変換 </summary>
        /// <param name="s">
        ///     変換する文字列</param>
        /// <returns>
        ///     変換後の文字列</returns>
        ///-------------------------------------------------------------
        public static string strSmallTolarge(string s)
        {
            s = s.Replace("ｬ", "ﾔ");
            s = s.Replace("ｭ", "ﾕ");
            s = s.Replace("ｮ", "ﾖ");
            s = s.Replace("ｯ", "ﾂ");

            return s;
        }

        //--------------------------------------------------------------
        /// <summary>
        ///     郵便番号住所で住所文字列を更新する </summary>
        /// <param name="st1">
        ///     住所文字列</param>
        /// <param name="st2">
        ///     郵便番号住所</param>
        /// <returns>
        ///     更新後文字列</returns>
        //--------------------------------------------------------------
        public static string addressUpdate(string st1, string st2)
        {
            st1 = Utility.strSmallTolarge(st1).Replace(" ", "").Trim();
            st2 = Utility.strSmallTolarge(st2).Replace(" ", "").Trim();

            if (st2 != "")
            {
                return st2 + " " + st1.Replace(st2, "");
            }
            else
            {
                return st1;
            }
        }

        //--------------------------------------------------------------------------------
        //  1バイト文字で構成された文字列であるか判定
        //
        //  1バイト文字のみで構成された文字列 : True
        //  2バイト文字が含まれている文字列   : False
        //--------------------------------------------------------------------------------
        public static bool isOneByteChar(string str)
        {
            byte[] byte_data = System.Text.Encoding.GetEncoding(932).GetBytes(str);
            if (byte_data.Length == str.Length)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     自らのロックファイルが存在したら削除する </summary>
        /// <param name="fPath">
        ///     パス</param>
        /// <param name="PcK">
        ///     自分のロックファイル文字列</param>
        ///-------------------------------------------------------------------------
        public static void deleteLockFile(string fPath, string PcK)
        {
            string FileName = fPath + global.LOCK_FILEHEAD + PcK + ".loc";

            if (System.IO.File.Exists(FileName))
            {
                System.IO.File.Delete(FileName);
            }
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     データフォルダにロックファイルが存在するか調べる </summary>
        /// <param name="fPath">
        ///     データフォルダパス</param>
        /// <returns>
        ///     true:ロックファイルあり、false:ロックファイルなし</returns>
        ///-------------------------------------------------------------------------
        public static Boolean existsLockFile(string fPath)
        {
            int s = System.IO.Directory.GetFiles(fPath, global.LOCK_FILEHEAD + "*.*", System.IO.SearchOption.TopDirectoryOnly).Count();

            if (s == 0)
            {
                return false; //LOCKファイルが存在しない
            }
            else
            {
                return true;   //存在する
            }
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     ロックファイルを登録する </summary>
        /// <param name="fPath">
        ///     書き込み先フォルダパス</param>
        /// <param name="LocName">
        ///     ファイル名</param>
        ///----------------------------------------------------------------
        public static void makeLockFile(string fPath, string LocName)
        {
            string FileName = fPath + global.LOCK_FILEHEAD + LocName + ".loc";

            //存在する場合は、処理なし
            if (System.IO.File.Exists(FileName))
            {
                return;
            }

            // ロックファイルを登録する
            try
            {
                System.IO.StreamWriter outFile = new System.IO.StreamWriter(FileName, false, System.Text.Encoding.GetEncoding(932));
                outFile.Close();
            }
            catch
            {
            }

            return;
        }

        ///-------------------------------------------------------------------------------
        /// <summary>
        ///     指定したファイルをロックせずに、System.Drawing.Imageを作成する。</summary>
        /// <param name="filename">
        ///     作成元のファイルのパス</param>
        /// <returns>
        ///     作成したSystem.Drawing.Image。</returns>
        ///-------------------------------------------------------------------------------
        public static System.Drawing.Image CreateImage(string filename)
        {
            System.IO.FileStream fs = new System.IO.FileStream(
                filename,
                System.IO.FileMode.Open,
                System.IO.FileAccess.Read);

            System.Drawing.Image img = System.Drawing.Image.FromStream(fs);

            fs.Close();
            return img;
        }

    }
}
