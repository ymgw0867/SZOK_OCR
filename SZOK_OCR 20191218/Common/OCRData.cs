using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace SZOK_OCR.Common
{
    class OCRData
    {
        #region エラー項目番号プロパティ
        //---------------------------------------------------
        //          エラー情報
        //---------------------------------------------------
        /// <summary>
        ///     エラーヘッダ行RowIndex</summary>
        public int _errHeaderIndex { get; set; }

        /// <summary>
        ///     エラー項目番号</summary>
        public int _errNumber { get; set; }

        /// <summary>
        ///     エラー明細行RowIndex </summary>
        public int _errRow { get; set; }

        /// <summary> 
        ///     エラーメッセージ </summary>
        public string _errMsg { get; set; }

        /// <summary> 
        ///     エラーなし </summary>
        public int eNothing = 0;

        /// <summary>
        ///     エラー項目 = 登録番号 </summary>
        public int eTourokuNum = 1;

        /// <summary>
        ///     エラー項目 = 車体番号 </summary>
        public int eShataiNum = 2;

        /// <summary> 
        ///     エラー項目 = 対象年月 </summary>
        public int eYear = 3;

        /// <summary> 
        ///     エラー項目 = 対象月 </summary>
        public int eMonth = 4;

        /// <summary> 
        ///     エラー項目 = 対象日 </summary>
        public int eDay = 5;
        
        /// <summary> 
        ///     エラー項目 = メーカー </summary>
        public int eMaker = 7;

        /// <summary> 
        ///     エラー項目 = 塗色 </summary>
        public int eColor = 8;

        /// <summary> 
        ///     エラー項目 = 車種 </summary>
        public int eStyle = 9;

        /// <summary> 
        ///     エラー項目 = 郵便番号 </summary>
        public int eZip1 = 10;

        /// <summary> 
        ///     エラー項目 = 郵便番号 </summary>
        public int eZip2 = 11;

        /// <summary> 
        ///     エラー項目 = 住所フリガナ </summary>
        public int eAddFuri = 12;

        /// <summary> 
        ///     エラー項目 = 住所 </summary>
        public int eAdd = 13;

        /// <summary> 
        ///     エラー項目 = フリガナ </summary>
        public int eFuri = 14;

        /// <summary> 
        ///     エラー項目 = 氏名 </summary>
        public int eName = 15;

        /// <summary> 
        ///     エラー項目 = TEL/携帯 </summary>
        public int eTel = 16;

        /// <summary> 
        ///     エラー項目 = 確認チェック </summary>
        public int eCheck = 17;

        /// <summary> 
        ///     エラー項目 = 車名 </summary>
        public int eCarName = 18;

        #endregion

        #region 時間チェック記号定数
        private const string cHOUR = "H";   // 時間をチェック
        private const string cMINUTE = "M"; // 分をチェック
        private const string cTIME = "HM";  // 時間・分をチェック
        #endregion
        
        ///-----------------------------------------------------------------------
        /// <summary>
        ///     ＣＳＶデータをＭＤＢに登録する：DataSet Version </summary>
        /// <param name="_InPath">
        ///     CSVデータパス</param>
        /// <param name="frmP">
        ///     プログレスバーフォームオブジェクト</param>
        /// <param name="dts">
        ///     データセット</param>
        ///-----------------------------------------------------------------------
        public void CsvToMdb(string _InPath, frmPrg frmP, szokDataSet dts)
        {
            string headerKey = string.Empty;    // ヘッダキー
            string prnKBN = string.Empty;       // 申請書ID
            string sName = string.Empty;        // 社員名
            string sShozoku = string.Empty;     // 所属名
            string sShozokuCode = string.Empty; // 所属コード

            //SqlDataReader dr = null;

            // テーブルアダプタ
            szokDataSetTableAdapters.防犯カードTableAdapter adp = new szokDataSetTableAdapters.防犯カードTableAdapter();
            
            try
            {
                // ＯＣＲ防犯データセット読み込み
                adp.Fill(dts.防犯カード);

                // 対象CSVファイル数を取得
                int cLen = System.IO.Directory.GetFiles(_InPath, "*.csv").Count();

                // CSVファイルを１つづつ取得します
                int cCnt = 0;
                foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                {
                    //件数カウント
                    cCnt++;

                    //プログレスバー表示
                    frmP.Text = "OCR変換CSVデータロード中　" + cCnt.ToString() + "/" + cLen.ToString();
                    frmP.progressValue = cCnt * 100 / cLen;
                    frmP.ProgressStep();

                    // CSVファイルインポート
                    var s = System.IO.File.ReadAllLines(files, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // 追加用SZOKDataSet.OCR防犯Rowオブジェクトを作成する
                        dts.防犯カード.Add防犯カードRow(setOCRRecRow(dts, stCSV));
                    }
                }

                // データベースへ反映
                adp.Update(dts.防犯カード);

                //CSVファイルを削除する
                foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                {
                    System.IO.File.Delete(files);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票CSVインポート処理", MessageBoxButtons.OK);
            }
            finally
            {
            }
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     ＣＳＶデータをＭＤＢに登録する：DataSet Version </summary>
        /// <param name="_InPath">
        ///     CSVデータパス</param>
        /// <param name="frmP">
        ///     プログレスバーフォームオブジェクト</param>
        /// <param name="dts">
        ///     データセット</param>
        ///-----------------------------------------------------------------------
        public void CsvToMdb(string _InPath, frmPrg frmP, string sLabel, string sName)
        {
            // テーブルアダプタ
            cardDataSetTableAdapters.SCAN_DATATableAdapter adp = new cardDataSetTableAdapters.SCAN_DATATableAdapter();

            cardDataSet dts = new cardDataSet();

            try
            {
                // 対象CSVファイル数を取得
                int cLen = System.IO.Directory.GetFiles(_InPath, "*.csv").Count();

                // CSVファイルを１件づつ取得します
                int cCnt = 0;
                foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                {
                    //件数カウント
                    cCnt++;

                    //プログレスバー表示
                    frmP.Text = "OCR変換CSVデータロード中　" + cCnt.ToString() + "/" + cLen.ToString();
                    frmP.progressValue = cCnt * 100 / cLen;
                    frmP.ProgressStep();

                    // CSVファイルインポート
                    var s = System.IO.File.ReadAllLines(files, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // 追加用SZOKDataSet.OCR防犯Rowオブジェクトを作成する
                        dts.SCAN_DATA.AddSCAN_DATARow(setOCRRecRow(dts, stCSV, sLabel, sName));
                    }
                }

                // データベースへ反映
                adp.Update(dts.SCAN_DATA);

                //CSVファイルを削除する
                foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                {
                    System.IO.File.Delete(files);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票CSVインポート処理", MessageBoxButtons.OK);
            }
            finally
            {
            }
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     追加用SZOKDataSet.OCR防犯Rowオブジェクトを作成する</summary>
        /// <param name="stCSV">
        ///     CSV配列</param>
        /// <returns>
        ///     追加するSZOKDataSet.防犯カードRowオブジェクト</returns>
        ///---------------------------------------------------------------------------------
        private szokDataSet.防犯カードRow setOCRRecRow(szokDataSet tblSt, string[] stCSV)
        {
            szokDataSet.防犯カードRow r = tblSt.防犯カード.New防犯カードRow();

            r.画像名 = Utility.GetStringSubMax(stCSV[1], 21);
            r.登録番号 = Utility.GetStringSubMax(stCSV[2], 11);
            r.車体番号 = Utility.GetStringSubMax(stCSV[3], 20);

            r.登録年 = Utility.GetStringSubMax(stCSV[4], 2);
            r.登録月 = Utility.GetStringSubMax(stCSV[5], 2);
            r.登録日 = Utility.GetStringSubMax(stCSV[6], 2);

            r.メーカー = Utility.GetStringSubMax(stCSV[7], 10);
            r.塗色 = Utility.GetStringSubMax(stCSV[8], 6);
            r.車種 = Utility.StrtoInt(carStyleZeroEdit(Utility.GetStringSubMax(stCSV[9], 2)));
            r.郵便番号1 = Utility.GetStringSubMax(stCSV[10], 3);
            r.郵便番号2 = Utility.GetStringSubMax(stCSV[11], 4);
            r.住所1 = Utility.GetStringSubMax(stCSV[12] + stCSV[13], 40).Trim();
            r.氏名 = Utility.GetStringSubMax(stCSV[14], 16);
            r.TEL携帯 = Utility.GetStringSubMax(stCSV[15], 4);
            r.TEL携帯2 = Utility.GetStringSubMax(stCSV[16], 4);
            r.TEL携帯3 = Utility.GetStringSubMax(stCSV[17], 4);
            r.更新年月日 = DateTime.Now;

            r.データ区分 = getDataKbn(stCSV[18]);
            r.車両番号1 = string.Empty;
            r.車両番号2 = string.Empty;
            r.車名 = string.Empty;
            r.住所漢字 = string.Empty;
            r.PC名 = string.Empty;
            r.CSV作成日 = string.Empty;
            r.備考 = string.Empty;
            r.更新年月日 = DateTime.Now;
            r.確認 = global.flgOff;       // 2016/01/25

            return r;
        }


        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     追加用cardDataSet.SCAN_DATARowオブジェクトを作成する</summary>
        /// <param name="stCSV">
        ///     CSV配列</param>
        /// <returns>
        ///     追加するcardDataSet.SCAN_DATARowオブジェクト</returns>
        ///---------------------------------------------------------------------------------
        private cardDataSet.SCAN_DATARow setOCRRecRow(cardDataSet dts, string[] stCSV, string sLabel, string sName)
        {
            cardDataSet.SCAN_DATARow r = dts.SCAN_DATA.NewSCAN_DATARow();

            r.画像名 = Utility.GetStringSubMax(stCSV[1], 21);
            r.登録番号 = Utility.GetStringSubMax(stCSV[2], 11);
            r.車体番号 = Utility.GetStringSubMax(stCSV[3], 20);

            r.登録年 = Utility.GetStringSubMax(stCSV[4], 2);
            r.登録月 = Utility.GetStringSubMax(stCSV[5], 2);
            r.登録日 = Utility.GetStringSubMax(stCSV[6], 2);

            r.メーカー = Utility.GetStringSubMax(stCSV[7], 10);
            r.塗色 = Utility.GetStringSubMax(stCSV[8], 6);
            r.車種 = Utility.StrtoInt(carStyleZeroEdit(Utility.GetStringSubMax(stCSV[9], 2)));
            r.郵便番号1 = Utility.GetStringSubMax(stCSV[10], 3);
            r.郵便番号2 = Utility.GetStringSubMax(stCSV[11], 4);
            r.住所1 = Utility.GetStringSubMax(stCSV[12] + stCSV[13], 40).Trim();
            r.住所2 = string.Empty;   // 2019/11/19
            r.氏名 = Utility.GetStringSubMax(stCSV[14], 16);
            r.TEL携帯 = Utility.GetStringSubMax(stCSV[15], 4);
            r.TEL携帯2 = Utility.GetStringSubMax(stCSV[16], 4);
            r.TEL携帯3 = Utility.GetStringSubMax(stCSV[17], 4);
            r.更新年月日 = DateTime.Now;

            r.データ区分 = getDataKbn(stCSV[18]);
            r.車両番号1 = string.Empty;
            r.車両番号2 = string.Empty;
            r.車名 = string.Empty;
            r.住所漢字 = string.Empty;
            r.PC名 = string.Empty;
            r.CSV作成日 = string.Empty;
            r.備考 = string.Empty;
            r.更新年月日 = DateTime.Now;
            r.ラベル  = sLabel;
            r.処理担当者 = sName;

            return r;
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     車種コード編集 </summary>
        /// <param name="s">
        ///     車種コード</param>
        /// <returns>
        ///     変換後車種コード</returns>
        ///-------------------------------------------------------------
        private string carStyleZeroEdit(string s)
        {
            string _s = string.Empty;

            // ハイフンを除去 2017/09/24
            s = s.Replace("-", "");

            if (Utility.StrtoInt(s.Trim()) == 9)
            {
                _s = s.Replace(" ", global.FLGOFF);
            }
            else
            {
                _s = s.Replace(" ", "");
            }
            
            return _s;
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     防犯登録カードの種類を判別します </summary>
        /// <param name="s">
        ///     文字列</param>
        /// <returns>
        ///     0:自転車, 1:原付</returns>
        ///--------------------------------------------------------
        private int getDataKbn(string s)
        {
            int rtn = 0;

            // 全てを自転車とする 2016/06/28

            //if (s.Contains("自") || s.Contains("転") || s.Contains("車"))
            //{
            //    // 自転車と判断
            //    rtn = 0;
            //}
            //else
            //{
            //    // 原付と判断
            //    rtn = 1;
            //}

            return rtn;
        }

        ///----------------------------------------------------------------------------------------
        /// <summary>
        ///     値1がemptyで値2がNot string.Empty のとき "0"を返す。そうではないとき値1をそのまま返す</summary>
        /// <param name="str1">
        ///     値1：文字列</param>
        /// <param name="str2">
        ///     値2：文字列</param>
        /// <returns>
        ///     文字列</returns>
        ///----------------------------------------------------------------------------------------
        private string hmStrToZero(string str1, string str2)
        {
            string rVal = str1;
            if (str1 == string.Empty && str2 != string.Empty)
                rVal = "0";

            return rVal;
        }

        ///--------------------------------------------------------------------------------------------------
        /// <summary>
        ///     エラーチェックメイン処理。
        ///     エラーのときOCRDataクラスのヘッダ行インデックス、フィールド番号、明細行インデックス、
        ///     エラーメッセージが記録される </summary>
        /// <param name="sIx">
        ///     開始ヘッダ行インデックス</param>
        /// <param name="eIx">
        ///     終了ヘッダ行インデックス</param>
        /// <param name="frm">
        ///     親フォーム</param>
        /// <param name="dts">
        ///     データセット</param>
        /// <param name="cID">
        ///     キー配列</param>
        /// <param name="z">
        ///     郵便番号CSV配列</param>
        /// <returns>
        ///     True:エラーなし、false:エラーあり</returns>
        ///-----------------------------------------------------------------------------------------------
        public Boolean errCheckMain(int sIx, int eIx, Form frm, szokDataSet dts, int [] cID, string [] z)
        {
            int rCnt = 0;

            // オーナーフォームを無効にする
            frm.Enabled = false;

            // プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = frm;
            frmP.Show();

            // レコード件数取得
            int cTotal = cID.Length;

            // 防犯登録カードデータ読み出し
            Boolean eCheck = true;

            for (int i = 0; i < cTotal; i++)
            {
                // データ件数加算
                rCnt++;

                //プログレスバー表示
                frmP.Text = "エラーチェック実行中　" + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt * 100 / cTotal;
                frmP.ProgressStep();

                // 指定範囲ならエラーチェックを実施する：（i:行index）
                if (i >= sIx && i <= eIx)
                {
                    // 防犯登録カードのコレクションを取得します
                    szokDataSet.防犯カードRow r = dts.防犯カード.Single(a => a.ID == cID[i]);

                    // エラーチェック実施
                    eCheck = errCheckData(dts, r, z);

                    if (!eCheck)　//エラーがあったとき
                    {
                        _errHeaderIndex = i;     // エラーとなったヘッダRowIndex
                        break;
                    }
                }
            }

            // いったんオーナーをアクティブにする
            frm.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            frm.Enabled = true;

            return eCheck;
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     エラー情報を取得します </summary>
        /// <param name="eID">
        ///     エラーデータのID</param>
        /// <param name="eNo">
        ///     エラー項目番号</param>
        /// <param name="eRow">
        ///     エラー明細行</param>
        /// <param name="eMsg">
        ///     表示メッセージ</param>
        ///---------------------------------------------------------------------------------
        private void setErrStatus(int eNo, int eRow, string eMsg)
        {
            //errHeaderIndex = eHRow;
            _errNumber = eNo;
            _errRow = eRow;
            _errMsg = eMsg;
        }

        ///-----------------------------------------------------------------------------------------------
        /// <summary>
        ///     項目別エラーチェック。
        ///     エラーのときヘッダ行インデックス、フィールド番号、明細行インデックス、エラーメッセージが記録される </summary>
        /// <param name="dts">
        ///     データセット</param>
        /// <param name="r">
        ///     防犯カード行コレクション</param>
        /// <param name="z">
        ///     郵便番号CSV配列</param>
        /// <returns>
        ///     エラーなし：true, エラー有り：false</returns>
        ///-----------------------------------------------------------------------------------------------
        /// 
        public Boolean errCheckData(szokDataSet dts, szokDataSet.防犯カードRow r, string [] z)
        {
            //string sDate;
            //DateTime eDate;

            // 登録番号
            if (r.登録番号 == string.Empty || r.登録番号.Length < 10)
            {
                setErrStatus(eTourokuNum, 0, "登録番号が正しくありません");
                return false;
            }

            if (r.登録番号.Substring(0, 3) != global.DATA_CPA)
            {
                setErrStatus(eTourokuNum, 0, "登録番号が正しくありません");
                return false;
            }

            // 登録番号登録済みチェック
            cardDataSet cDts = new cardDataSet();
            cardDataSetTableAdapters.防犯登録データTableAdapter cAdp = new cardDataSetTableAdapters.防犯登録データTableAdapter();
            
            //// 2019/06/25 コメント化
            //cAdp.Fill(cDts.防犯登録データ);

            //if (cDts.防犯登録データ.Any(a => (a.Is除外Null() || a.除外 == global.flgOff) && a.登録番号 == r.登録番号))
            //{
            //    setErrStatus(eTourokuNum, 0, "過去に登録されている登録番号です");
            //    return false;
            //}

            // Fillで絞込 2019/06/25
            cAdp.FillByISNumber(cDts.防犯登録データ, r.登録番号);

            if (cDts.防犯登録データ.Count() > 0)
            {
                setErrStatus(eTourokuNum, 0, "過去に登録されている登録番号です");
                return false;
            }

            // 登録番号重複チェック
            if (dts.防犯カード.Any(a => a.ID != r.ID && a.登録番号 == r.登録番号))
            {
                setErrStatus(eTourokuNum, 0, "現在、読み込み中データに同じ登録番号が複数あります");
                return false;
            }
            
            // 車体番号
            if (r.車体番号 == string.Empty)
            {
                setErrStatus(eShataiNum, 0, "車体番号が未入力です");
                return false;
            }

            // 登録年
            if (!Utility.NumericCheck(r.登録年.ToString()))
            {
                setErrStatus(eYear, 0, "登録年が正しくありません");
                return false;
            }

            // 登録月
            if (!Utility.NumericCheck(r.登録月.ToString()))
            {
                setErrStatus(eMonth, 0, "月が正しくありません");
                return false;
            }

            if (int.Parse(r.登録月.ToString()) < 1 || int.Parse(r.登録月.ToString()) > 12)
            {
                setErrStatus(eMonth, 0, "月が正しくありません");
                return false;
            }

            // 登録日
            if (!Utility.NumericCheck(r.登録日.ToString()))
            {
                setErrStatus(eDay, 0, "月が正しくありません");
                return false;
            }

            if (int.Parse(r.登録日.ToString()) < 1 || int.Parse(r.登録日.ToString()) > 31)
            {
                setErrStatus(eDay, 0, "登録日が正しくありません");
                return false;
            }

            // 存在する日付・翌日以降チェック
            DateTime dt;
            string sDt = "20" + r.登録年.ToString() + "/" + r.登録月.ToString() + "/" + r.登録日.ToString();
            if (!DateTime.TryParse(sDt, out dt))
            {
                setErrStatus(eYear, 0, "登録年月日が正しくありません");
                return false;
            }
            else
            {
                if (dt > DateTime.Today)
                {
                    setErrStatus(eYear, 0, "登録年月日が翌日以降になっています");
                    return false;
                }
            }

            // メーカー・漢字が含まれていたらエラー 2016/07/15
            if (!Utility.isOneByteChar(r.メーカー))
            {
                setErrStatus(eMaker, 0, "漢字が含まれています");
                return false;
            }
            
            // カラー・漢字が含まれていたらエラー 2016/07/15
            if (!Utility.isOneByteChar(r.塗色))
            {
                setErrStatus(eColor, 0, "漢字が含まれています");
                return false;
            }

            // 車種
            bool shashu = false;
            string dArray = string.Empty;

            global g = new global();
            for (int i = 0; i < g.arrStyle.GetLength(0); i++)
            {
                if (g.arrStyle[i, 0] == r.車種.ToString().PadLeft(2, '0'))
                {
                    shashu = true;
                    dArray = g.arrStyle[i, 2];

                    break;
                }
            }

            if (!shashu)
            {
                setErrStatus(eStyle, 0, "車種が正しくありません");
                return false;
            }
            else
            {
                // 自転車・原付の車種か？
                if (r.データ区分 != 90 && dArray != r.データ区分.ToString())
                {
                    setErrStatus(eStyle, 0, "カード種類と車種が一致していません");
                    return false;
                }
            }

            // 車名・漢字が含まれていたらエラー 2016/07/15
            if (!Utility.isOneByteChar(r.車名))
            {
                setErrStatus(eCarName, 0, "漢字が含まれています");
                return false;
            }

            // 郵便番号
            // 未入力のとき
            if (r.郵便番号1 == string.Empty || r.郵便番号2 == string.Empty)
            {
                setErrStatus(eZip1, 0, "郵便番号が正しくありません");
                return false;
            }

            string [] zipAdd = null;
            int iZ = 0;

            // 郵便番号CSVに該当があるか（住所に"ｹﾝｶﾞｲ"と記入されているときはエラーとしない）
            if (!r.住所1.Contains(global.KENGAI_ADD))
            {
                bool zipStatus = false;
                string zipText = r.郵便番号1 + r.郵便番号2;

                if (zipText.Length == 7)
                {
                    if (z != null)
                    {
                        foreach (var t in z)
                        {
                            string[] zip = t.Split(',');

                            if (zipText == zip[2].Replace("\"", ""))
                            {
                                // 「同一郵便番号で複数の地名」への対応：該当地名を配列に格納 2016/06/07
                                Array.Resize(ref zipAdd, iZ + 1);
                                zipAdd[iZ] = (zip[4] + " " + zip[5]).Replace("\"", "");
                                zipStatus = true;
                                iZ++;
                            }
                        }

                        // 郵便番号該当なし
                        if (!zipStatus)
                        {
                            setErrStatus(eZip1, 0, "郵便番号が正しくありません");
                            return false;
                        }
                    }
                }
                else
                {
                    setErrStatus(eZip1, 0, "郵便番号が正しくありません");
                    return false;
                }
            }

            // 住所
            string sAddress = r.住所1;

            if (sAddress.Trim() == string.Empty)
            {
                setErrStatus(eAddFuri, 0, "住所が未入力です");
                return false;
            }

            // 漢字が含まれていたらエラー 2016/07/15
            if (!Utility.isOneByteChar(sAddress))
            {
                setErrStatus(eAddFuri, 0, "漢字が含まれています");
                return false;
            }

            // 住所と郵便番号が一致しているか？
            string str1 = Utility.strSmallTolarge(Utility.getStrConv(sAddress.Replace(" ", "").Trim()));
            bool errAddStatus = false; 

            // 住所配列から判断する：2016/06/07
            if (zipAdd != null)
            {
                for (int i = 0; i < zipAdd.Length; i++)
                {
                    // 以下に掲載がない場合の文字列を除去する：2016/06/07
                    string str2 = Utility.strSmallTolarge((Utility.getStrConv(zipAdd[i].Replace(global.IKAKEISAI_ADD, "").Replace(" ", "").Trim())));

                    //  ()表記は除去する 2016/06/07
                    int zC = str2.IndexOf("(");

                    if (zC != -1)
                    {
                        str2 = str2.Replace(str2.Substring(zC, str2.Length - zC), "");
                    }
                    
                    if (str1.Contains(str2))
                    {
                        errAddStatus = true;
                        break;
                    }
                }

                if (!errAddStatus)
                {
                    setErrStatus(eAddFuri, 0, "郵便番号と住所が一致していません");
                    return false;
                }
            }

            // 氏名（フリガナ
            if (r.氏名.Trim() == string.Empty)
            {
                setErrStatus(eFuri, 0, "氏名が未入力です");
                return false;
            }

            // 氏名（フリガナ）・漢字が含まれていたらエラー 2016/07/15
            if (!Utility.isOneByteChar(r.氏名))
            {
                setErrStatus(eFuri, 0, "漢字が含まれています");
                return false;
            }

            // TEL/携帯番号
            string sTel = r.TEL携帯 + "-" + r.TEL携帯2 + "-" + r.TEL携帯3;

            if (sTel.Replace("-", string.Empty) == string.Empty)
            {
                setErrStatus(eTel, 0, "TEL/携帯番号が未入力です");
                return false;
            }

            // TEL/携帯番号：正規化チェック
            if (!Utility.regexTelNum(sTel))
            {
                setErrStatus(eTel, 0, "TEL/携帯番号が正しくありません");
                return false;
            }

            // 確認チェック
            if (r.Is確認Null() || r.確認 == global.flgOff)
            {
                setErrStatus(eCheck, 0, "画面確認が未チェックです");
                return false;
            }

            return true;
        }
    }
}
