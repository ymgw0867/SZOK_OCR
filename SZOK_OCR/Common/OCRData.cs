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
        public int eYearMonth = 3;

        /// <summary> 
        ///     エラー項目 = 対象月 </summary>
        public int eMonth = 4;

        /// <summary> 
        ///     エラー項目 = 対象日 </summary>
        public int eDay = 5;

        /// <summary> 
        ///     エラー項目 = 警察番号 </summary>
        public int eKeisatsuNum = 6;

        /// <summary> 
        ///     エラー項目 = 警察整理番号 </summary>
        public int eSeiriNum = 7;

        /// <summary> 
        ///     エラー項目 = メーカー </summary>
        public int eMaker = 8;

        /// <summary> 
        ///     エラー項目 = 塗色 </summary>
        public int eColor = 9;

        /// <summary> 
        ///     エラー項目 = 車種 </summary>
        public int eStyle = 10;

        /// <summary> 
        ///     エラー項目 = 郵便番号 </summary>
        public int eZip1 = 11;

        /// <summary> 
        ///     エラー項目 = 郵便番号 </summary>
        public int eZip2 = 12;

        /// <summary> 
        ///     エラー項目 = 住所フリガナ </summary>
        public int eAddFuri = 13;

        /// <summary> 
        ///     エラー項目 = 住所 </summary>
        public int eAdd = 14;

        /// <summary> 
        ///     エラー項目 = フリガナ </summary>
        public int eFuri = 15;

        /// <summary> 
        ///     エラー項目 = 氏名 </summary>
        public int eName = 16;

        /// <summary> 
        ///     エラー項目 = TEL/携帯 </summary>
        public int eTel = 17;
        
        #endregion

        #region フィールド定義
        /// <summary> 
        ///     警告項目 = 時間外1.25時 </summary>
        public int [] wZ125HM = new int[global.MAX_GYO];

        /// <summary> 
        ///     実働時間 </summary>
        public double _workTime;

        /// <summary> 
        ///     深夜稼働時間 </summary>
        public double _workShinyaTime;
        #endregion

        #region 単位時間フィールド
        /// <summary> 
        ///     ３０分単位 </summary>
        private int tanMin30 = 30;

        /// <summary> 
        ///     １５分単位 </summary> 
        private int tanMin15 = 15;

        /// <summary> 
        ///     １分単位 </summary>
        private int tanMin1 = 1;
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
        /// <param name="dbName">
        ///     会社領域データベース名</param>
        /// <param name="comName">
        ///     会社名</param>
        /// <param name="comNo">
        ///     会社領域会社番号</param>
        ///-----------------------------------------------------------------------
        public void CsvToMdb(string _InPath, frmPrg frmP, SZOKDataSet dts)
        {
            string headerKey = string.Empty;    // ヘッダキー
            string prnKBN = string.Empty;       // 申請書ID
            string sName = string.Empty;        // 社員名
            string sShozoku = string.Empty;     // 所属名
            string sShozokuCode = string.Empty; // 所属コード

            SqlDataReader dr = null;

            // テーブルアダプタ
            SZOKDataSetTableAdapters.OCR防犯TableAdapter adp = new SZOKDataSetTableAdapters.OCR防犯TableAdapter();

            // テーブルセットオブジェクト
            SZOKDataSet tblSt = new SZOKDataSet();

            try
            {
                // ＯＣＲ防犯データセット読み込み
                adp.Fill(tblSt.OCR防犯);

                // 対象CSVファイル数を取得
                int cLen = System.IO.Directory.GetFiles(_InPath, "*.csv").Count();

                //string [] t = System.IO.Directory.GetFiles(_InPath, "*.csv");
                //int cLen = t.Length;

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
                        tblSt.OCR防犯.AddOCR防犯Row(setOCRRecRow(tblSt, stCSV));
                    }
                }

                // データベースへ反映
                adp.Update(tblSt);

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
        ///     追加するSZOKDataSet.OCR防犯Rowオブジェクト</returns>
        ///---------------------------------------------------------------------------------
        private SZOKDataSet.OCR防犯Row setOCRRecRow(SZOKDataSet tblSt, string[] stCSV)
        {
            SZOKDataSet.OCR防犯Row r = tblSt.OCR防犯.NewOCR防犯Row();

            r.画像名 = Utility.GetStringSubMax(stCSV[1], 21);

            r.登録年 = Utility.GetStringSubMax(stCSV[2], 2);
            r.登録月 = Utility.GetStringSubMax(stCSV[3], 2);
            r.登録日 = Utility.GetStringSubMax(stCSV[4], 2);
            r.登録番号 = Utility.GetStringSubMax(stCSV[5], 12);
            r.車体番号 = Utility.GetStringSubMax(stCSV[6], 15);
            r.警察整理番号 = Utility.GetStringSubMax(stCSV[7], 1);
            r.整理番号 = Utility.GetStringSubMax(stCSV[8], 10);
            r.メーカー = Utility.GetStringSubMax(stCSV[9], 1);
            r.塗色 = Utility.GetStringSubMax(stCSV[10], 1);
            r.車種 = Utility.GetStringSubMax(stCSV[11], 1);
            r.郵便番号 = Utility.GetStringSubMax(stCSV[12], 1) + Utility.GetStringSubMax(stCSV[13], 1) + Utility.GetStringSubMax(stCSV[14], 1) +
                        Utility.GetStringSubMax(stCSV[15], 1) + Utility.GetStringSubMax(stCSV[16], 1) + Utility.GetStringSubMax(stCSV[17], 1) +
                        Utility.GetStringSubMax(stCSV[18], 1);
            r.フリガナ = Utility.GetStringSubMax(stCSV[19], 15);
            r.氏名 = Utility.GetStringSubMax(stCSV[20], 10);
            r.TEL携帯1 = Utility.GetStringSubMax(stCSV[21], 4);
            r.TEL携帯2 = Utility.GetStringSubMax(stCSV[22], 4);
            r.TEL携帯3 = Utility.GetStringSubMax(stCSV[23], 4);
            r.更新年月日 = DateTime.Now;

            return r;
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

        /////--------------------------------------------------------------------------------------------------
        ///// <summary>
        /////     エラーチェックメイン処理。
        /////     エラーのときOCRDataクラスのヘッダ行インデックス、フィールド番号、明細行インデックス、
        /////     エラーメッセージが記録される </summary>
        ///// <param name="sIx">
        /////     開始ヘッダ行インデックス</param>
        ///// <param name="eIx">
        /////     終了ヘッダ行インデックス</param>
        ///// <param name="frm">
        /////     親フォーム</param>
        ///// <param name="dts">
        /////     データセット</param>
        ///// <param name="dbName">
        /////     給与大臣データベース番号</param>
        ///// <param name="comNo">
        /////     給与大臣会社番号</param>
        ///// <returns>
        /////     True:エラーなし、false:エラーあり</returns>
        /////-----------------------------------------------------------------------------------------------
        //public Boolean errCheckMain(int sIx, int eIx, Form frm, SZOKDataSet dts, string dbName, string comNo)
        //{
        //    int rCnt = 0;

        //    // オーナーフォームを無効にする
        //    frm.Enabled = false;

        //    // プログレスバーを表示する
        //    frmPrg frmP = new frmPrg();
        //    frmP.Owner = frm;
        //    frmP.Show();

        //    // レコード件数取得
        //    int cTotal = dts.OCR防犯.Rows.Count;

        //    // 出勤簿データ読み出し
        //    Boolean eCheck = true;

        //    for (int i = 0; i < cTotal; i++)
        //    {
        //        //データ件数加算
        //        rCnt++;

        //        //プログレスバー表示
        //        frmP.Text = "エラーチェック実行中　" + rCnt.ToString() + "/" + cTotal.ToString();
        //        frmP.progressValue = rCnt * 100 / cTotal;
        //        frmP.ProgressStep();

        //        //指定範囲ならエラーチェックを実施する：（i:行index）
        //        if (i >= sIx && i <= eIx)
        //        {
        //            // 勤務票ヘッダ行のコレクションを取得します
        //            SZOKDataSet.OCR防犯Row r = (SZOKDataSet.OCR防犯Row)dts.OCR防犯.Rows[i];

        //            // エラーチェック実施
        //            eCheck = errCheckData(dts, r, dbName, comNo);

        //            if (!eCheck)　//エラーがあったとき
        //            {
        //                _errHeaderIndex = i;     // エラーとなったヘッダRowIndex
        //                break;
        //            }
        //        }
        //    }

        //    // いったんオーナーをアクティブにする
        //    frm.Activate();

        //    // 進行状況ダイアログを閉じる
        //    frmP.Close();

        //    // オーナーのフォームを有効に戻す
        //    frm.Enabled = true;

        //    return eCheck;
        //}

        /////---------------------------------------------------------------------------------
        ///// <summary>
        /////     エラー情報を取得します </summary>
        ///// <param name="eID">
        /////     エラーデータのID</param>
        ///// <param name="eNo">
        /////     エラー項目番号</param>
        ///// <param name="eRow">
        /////     エラー明細行</param>
        ///// <param name="eMsg">
        /////     表示メッセージ</param>
        /////---------------------------------------------------------------------------------
        //private void setErrStatus(int eNo, int eRow, string eMsg)
        //{
        //    //errHeaderIndex = eHRow;
        //    _errNumber = eNo;
        //    _errRow = eRow;
        //    _errMsg = eMsg;
        //}

        /////-----------------------------------------------------------------------------------------------
        ///// <summary>
        /////     項目別エラーチェック。
        /////     エラーのときヘッダ行インデックス、フィールド番号、明細行インデックス、エラーメッセージが記録される </summary>
        ///// <param name="dts">
        /////     データセット</param>
        ///// <param name="r">
        /////     勤務票ヘッダ行コレクション</param>
        ///// <param name="dbName">
        /////     給与大臣データベース番号</param>
        ///// <param name="comNo">
        /////     給与大臣会社番号</param>
        ///// <returns>
        /////     エラーなし：true, エラー有り：false</returns>
        /////-----------------------------------------------------------------------------------------------
        ///// 
        //public Boolean errCheckData(SZOKDataSet dts, SZOKDataSet.OCR防犯Row r, string dbName, string comNo)
        //{
        //    string sDate;
        //    DateTime eDate;

        //    // 勤怠区分データテーブルを取得します
        //    JSDataSetTableAdapters.勤怠区分TableAdapter adp_k = new JSDataSetTableAdapters.勤怠区分TableAdapter();
        //    adpMn.勤怠区分TableAdapter = adp_k;
        //    adpMn.勤怠区分TableAdapter.Fill(dts.勤怠区分);

        //    // 特別手当区分データテーブルを取得します
        //    JSDataSetTableAdapters.特別手当種類TableAdapter adp_T = new JSDataSetTableAdapters.特別手当種類TableAdapter();
        //    adpMn.特別手当種類TableAdapter = adp_T;
        //    adpMn.特別手当種類TableAdapter.Fill(dts.特別手当種類);

        //    // 帳票種別を取得
        //    string shain_Type = r.帳票番号.ToString();

        //    // 対象年
        //    if (Utility.NumericCheck(r.年.ToString()) == false)
        //    {
        //        setErrStatus(eYearMonth, 0, "年が正しくありません");
        //        return false;
        //    }

        //    if (r.年 < 1)
        //    {
        //        setErrStatus(eYearMonth, 0, "年が正しくありません");
        //        return false;
        //    }

        //    if (r.年 != global.cnfYear)
        //    {
        //        setErrStatus(eYearMonth, 0, "対象年（" + global.cnfYear + "年）と一致していません");
        //        return false;
        //    }

        //    // 対象月
        //    if (!Utility.NumericCheck(r.月.ToString()))
        //    {
        //        setErrStatus(eMonth, 0, "月が正しくありません");
        //        return false;
        //    }

        //    if (int.Parse(r.月.ToString()) < 1 || int.Parse(r.月.ToString()) > 12)
        //    {
        //        setErrStatus(eMonth, 0, "月が正しくありません");
        //        return false;
        //    }

        //    if (int.Parse(r.月.ToString()) != global.cnfMonth)
        //    {
        //        setErrStatus(eMonth, 0, "対象月（" + global.cnfMonth + "月）と一致していません");
        //        return false;
        //    }

        //    // 対象年月
        //    sDate = r.年.ToString() + "/" + r.月.ToString() + "/01";
        //    if (DateTime.TryParse(sDate, out eDate) == false)
        //    {
        //        setErrStatus(eYearMonth, 0, "年月が正しくありません");
        //        return false;
        //    }

        //    // 社員番号
        //    // 数字以外のとき
        //    if (!Utility.NumericCheck(Utility.NulltoStr(r.個人番号)))
        //    {
        //        setErrStatus(eShainNo, 0, "社員番号が入力されていません");
        //        return false;
        //    }

        //    // 社員、パートは給与大臣の社員情報を参照します
        //    if (r.帳票番号 == int.Parse(global.SHAIN_ID) || r.帳票番号 == int.Parse(global.PART_ID))
        //    {
        //        // 給与大臣の社員情報検索
        //        string Nm = string.Empty;
        //        string result = errCheckShainCode(dts, Utility.bldShainCode(r.個人番号.ToString()), out Nm);
             
        //        // マスター未登録
        //        if (result == global.NO_MASTER)
        //        {
        //            setErrStatus(eShainNo, 0, "マスター未登録の社員番号です");
        //            return false;
        //        }

        //        // 休職中社員
        //        if (result == global.NO_KYUSHOKU)
        //        {
        //            setErrStatus(eShainNo, 0, "休職中の社員です");
        //            return false;
        //        }

        //        // 退職社員
        //        if (result == global.NO_TAISHOKU)
        //        {
        //            setErrStatus(eShainNo, 0, "退職した社員です");
        //            return false;
        //        }
        //    }

        //    // 出向社員は出向社員データテーブルを参照します
        //    if (r.帳票番号 == int.Parse(global.SHUKKOU_ID))    
        //    {
        //        // 出向社員名簿データテーブルを参照戻り値：社員名
        //        string sName;

        //        // 出向社員名簿データテーブルを参照
        //        string result = getCheckShuShain(dts, r.個人番号.ToString().PadLeft(4, '0'), out sName);

        //        // マスター未登録
        //        if (result == global.NO_MASTER)
        //        {
        //            setErrStatus(eShainNo, 0, "出向社員名簿に未登録の社員番号です");
        //            return false;
        //        }

        //        // 在籍区分が０：非在籍扱い
        //        if (result == global.NO_ZAISEKI)
        //        {
        //            setErrStatus(eShainNo, 0, "該当する社員は現在、在籍していません");
        //            return false;
        //        }
        //    }

        //    // 同じ社員番号の勤務票データが複数存在しているか
        //    if (!getSameNumber(dts.勤務票ヘッダ, r.帳票番号, r.個人番号, r.ID))
        //    {
        //        setErrStatus(eShainNo, 0, "同じ帳票ID、社員番号のデータが複数あります");
        //        return false;
        //    }

        //    // 日付別データ
        //    int iX = 0;
        //    string k = string.Empty;    // 特別休暇記号
        //    string yk = string.Empty;   // 有給記号
        //    int stKbn = 0;              // 勤怠区分の出勤退勤区分
            
        //    // 勤務票明細データ行を取得
        //    var mData = dts.勤務票明細.Where(a => a.ヘッダID == r.ID).OrderBy(a => a.ID);

        //    foreach (var m in mData)
        //    {
        //        // 日付インデックス加算
        //        iX++;

        //        // 勤怠区分初期化
        //        stKbn = 0;

        //        // 日付は数字か
        //        if (!Utility.NumericCheck(m.日付))
        //        {
        //            setErrStatus(eDay, iX - 1, "日が正しくありません");
        //            return false;
        //        }

        //        sDate = r.年.ToString() + "/" + r.月.ToString() + "/" + m.日付.ToString();

        //        // 存在しない日付に記入があるとき
        //        if (!DateTime.TryParse(sDate, out eDate))
        //        {
        //            if (Utility.NulltoStr(m.勤怠区分) != string.Empty ||
        //            Utility.NulltoStr(m.開始時) != string.Empty || Utility.NulltoStr(m.開始分) != string.Empty || 
        //            Utility.NulltoStr(m.終了時) != string.Empty || Utility.NulltoStr(m.終了分) != string.Empty || 
        //            Utility.NulltoStr(m.時間外時) != string.Empty || Utility.NulltoStr(m.時間外分) != string.Empty || 
        //            Utility.NulltoStr(m.時間外12時) != string.Empty || Utility.NulltoStr(m.時間外12分) != string.Empty || 
        //            Utility.NulltoStr(m.時間外125時) != string.Empty || Utility.NulltoStr(m.時間外125分) != string.Empty || 
        //            Utility.NulltoStr(m.休日出勤時) != string.Empty || Utility.NulltoStr(m.休日出勤分) != string.Empty || 
        //            Utility.NulltoStr(m.深夜時) != string.Empty || Utility.NulltoStr(m.深夜分) != string.Empty || 
        //            m.呼出 != global.flgOff || m.交替 != global.flgOff || 
        //            Utility.NulltoStr(m.遅早退時) != string.Empty || Utility.NulltoStr(m.遅早退分) != string.Empty || 
        //            Utility.NulltoStr(m.所定内休日時) != string.Empty || Utility.NulltoStr(m.所定内休日分) != string.Empty || 
        //            Utility.NulltoStr(m.特別手当時) != string.Empty || Utility.NulltoStr(m.特別手当分) != string.Empty)
        //            {
        //                setErrStatus(eDay, iX - 1, "この行には記入できません");
        //                return false;
        //            }
        //        }

        //        // 無記入の行はチェック対象外とする
        //        if (Utility.NulltoStr(m.勤怠区分) == string.Empty &&
        //            Utility.NulltoStr(m.開始時) == string.Empty && Utility.NulltoStr(m.開始分) == string.Empty &&
        //            Utility.NulltoStr(m.終了時) == string.Empty && Utility.NulltoStr(m.終了分) == string.Empty &&
        //            Utility.NulltoStr(m.時間外時) == string.Empty && Utility.NulltoStr(m.時間外分) == string.Empty &&
        //            Utility.NulltoStr(m.時間外12時) == string.Empty && Utility.NulltoStr(m.時間外12分) == string.Empty &&
        //            Utility.NulltoStr(m.時間外125時) == string.Empty && Utility.NulltoStr(m.時間外125分) == string.Empty &&
        //            Utility.NulltoStr(m.休日出勤時) == string.Empty && Utility.NulltoStr(m.休日出勤分) == string.Empty &&
        //            Utility.NulltoStr(m.深夜時) == string.Empty && Utility.NulltoStr(m.深夜分) == string.Empty &&
        //            m.呼出 == global.flgOff && m.交替 == global.flgOff &&
        //            Utility.NulltoStr(m.遅早退時) == string.Empty && Utility.NulltoStr(m.遅早退分) == string.Empty &&
        //            Utility.NulltoStr(m.所定内休日時) == string.Empty && Utility.NulltoStr(m.所定内休日分) == string.Empty &&
        //            Utility.NulltoStr(m.特別手当時) == string.Empty && Utility.NulltoStr(m.特別手当分) == string.Empty)
        //        {
        //            continue;
        //        }

        //        // 勤怠区分チェック
        //        if (m.勤怠区分 != string.Empty)
        //        {
        //            if (!errCheckKintaiCode(m, dts, iX, r.帳票番号)) return false;

        //            // 勤怠区分マスターの出勤区分を取得
        //            JSDataSet.勤怠区分Row kR = dts.勤怠区分.FindByID(m.勤怠区分);
        //            if (kR != null) stKbn = kR.出勤区分;
        //        }
        //        else stKbn = global.flgOff;

        //        // 勤怠区分の出退勤区分が「0」で勤怠区分以外が無記入の行はチェック対象外とする
        //        if (stKbn == global.flgOff &&
        //            Utility.NulltoStr(m.開始時) == string.Empty && Utility.NulltoStr(m.開始分) == string.Empty &&
        //            Utility.NulltoStr(m.終了時) == string.Empty && Utility.NulltoStr(m.終了分) == string.Empty &&
        //            Utility.NulltoStr(m.時間外時) == string.Empty && Utility.NulltoStr(m.時間外分) == string.Empty &&
        //            Utility.NulltoStr(m.時間外12時) == string.Empty && Utility.NulltoStr(m.時間外12分) == string.Empty &&
        //            Utility.NulltoStr(m.時間外125時) == string.Empty && Utility.NulltoStr(m.時間外125分) == string.Empty &&
        //            Utility.NulltoStr(m.休日出勤時) == string.Empty && Utility.NulltoStr(m.休日出勤分) == string.Empty &&
        //            Utility.NulltoStr(m.深夜時) == string.Empty && Utility.NulltoStr(m.深夜分) == string.Empty &&
        //            m.呼出 == global.flgOff && m.交替 == global.flgOff &&
        //            Utility.NulltoStr(m.遅早退時) == string.Empty && Utility.NulltoStr(m.遅早退分) == string.Empty &&
        //            Utility.NulltoStr(m.所定内休日時) == string.Empty && Utility.NulltoStr(m.所定内休日分) == string.Empty &&
        //            Utility.NulltoStr(m.特別手当時) == string.Empty && Utility.NulltoStr(m.特別手当分) == string.Empty)
        //        {
        //            continue;
        //        }
                
        //        // 出勤時刻・退勤時刻チェック
        //        if (!errCheckTime(m, "出退時間", tanMin1, iX, stKbn)) return false;

        //        // 時間外チェック（※社員、出向）
        //        if (r.帳票番号.ToString() == global.SHAIN_ID || r.帳票番号.ToString() == global.SHUKKOU_ID)
        //        {
        //            if (!errCheckZan(m, "時間外", tanMin30, iX)) return false;
        //        }

        //        // 休日出勤チェック （※社員、出向）
        //        if (r.帳票番号.ToString() == global.SHAIN_ID || r.帳票番号.ToString() == global.SHUKKOU_ID)
        //        {
        //            if (!errCheckKyujitsuShukkin(m, "休日出勤", tanMin30, iX)) return false;
        //        }

        //        // 深夜勤務チェック （※社員、出向は30分単位）2014/03/28
        //        if (r.帳票番号.ToString() == global.SHAIN_ID || r.帳票番号.ToString() == global.SHUKKOU_ID)
        //        {
        //            if (!errCheckShinya(m, "深夜勤務", tanMin30, iX)) return false;
        //        }

        //        // 深夜勤務チェック （※パートは15分単位）2014/03/28
        //        if (r.帳票番号.ToString() == global.PART_ID)
        //        {
        //            if (!errCheckShinya(m, "深夜勤務", tanMin15, iX)) return false;
        //        }

        //        // 呼出区分チェック（※社員、出向）
        //        if (r.帳票番号.ToString() == global.SHAIN_ID || r.帳票番号.ToString() == global.SHUKKOU_ID)
        //        {
        //            if (!errCheckYCD(m, "呼出区分", iX)) return false;
        //        }

        //        // 交替区分チェック（※社員、出向）
        //        if (r.帳票番号.ToString() == global.SHAIN_ID || r.帳票番号.ToString() == global.SHUKKOU_ID)
        //        {
        //            if (!errCheckKotaiCD(m, "交替区分", iX)) return false;
        //        }

        //        // 遅刻早退外出チェック（※社員、出向）
        //        if (r.帳票番号.ToString() == global.SHAIN_ID || r.帳票番号.ToString() == global.SHUKKOU_ID)
        //        {
        //            if (!errCheckChikoku(m, "遅刻早退外出", tanMin30, iX)) return false;
        //        }

        //        // 所定内・休日時間チェック（※パート）
        //        if (r.帳票番号.ToString() == global.PART_ID)
        //        {
        //            //if (!errCheckShotei(m, "所定内・休日", tanMin30, iX)) return false;

        //            // 2014/05/07 「15分単位以外はエラー」に変更
        //            if (!errCheckShotei(m, "所定内・休日", tanMin15, iX)) return false;
        //        }

        //        // 時間外1.2チェック（※パート）
        //        if (r.帳票番号.ToString() == global.PART_ID)
        //        {
        //            if (!errCheckZan12(m, "時間外1.2", tanMin15, iX)) return false;
        //        }

        //        // 時間外1.25チェック（※パート）
        //        if (r.帳票番号.ToString() == global.PART_ID)
        //        {
        //            if (!errCheckZan125(m, "時間外1.25", tanMin15, iX)) return false;
        //        }

        //        // 特別手当チェック（※パート）
        //        if (r.帳票番号.ToString() == global.PART_ID)
        //        {
        //            if (!errCheckTokuTeate(m, "特別手当", tanMin15, iX)) return false;
        //        }
        //    }

        //    // 社員・出向社員のみ
        //    if (r.帳票番号.ToString() == global.SHAIN_ID || r.帳票番号.ToString() == global.SHUKKOU_ID)
        //    {
        //        // カット時間チェック（※社員、出向）
        //        if (r.帳票番号.ToString() == global.SHAIN_ID || r.帳票番号.ToString() == global.SHUKKOU_ID)
        //        {
        //            if (!errCheckCutTime(r, "カット時間", tanMin30)) return false;
        //        }

        //        // 特殊作業時間チェック（※社員、出向）
        //        if (r.帳票番号.ToString() == global.SHAIN_ID || r.帳票番号.ToString() == global.SHUKKOU_ID)
        //        {
        //            if (!errCheckTokushu(r, "特殊作業時間", tanMin30)) return false;
        //        }
        //    }

        //    // パートのみ
        //    if (r.帳票番号.ToString() == global.PART_ID)
        //    {
        //        // 特別手当区分チェック（※パート）
        //        if (r.帳票番号.ToString() == global.PART_ID)
        //        {
        //            if (!errCheckTteCD(dts, r, "特別手当区分")) return false;
        //        }
        //    }

        //    return true;
        //}

        //private bool errCheckKintaiCode(JSDataSet.勤務票明細Row m, JSDataSet dts, int iX, int cID)
        //{
        //    // 無記入は戻す
        //    if (Utility.NulltoStr(m.勤怠区分) == string.Empty) return true;

        //    // 勤怠区分のチェック
        //    JSDataSet.勤怠区分Row kR = dts.勤怠区分.FindByID(m.勤怠区分);

        //    if (kR == null) // マスター登録区分か検証する
        //    {
        //        setErrStatus(eKintaiKigou, iX - 1, "勤怠記号が正しくありません");
        //        return false;
        //    }
        //    else
        //    {
        //        if (cID == int.Parse(global.SHAIN_ID) || cID == int.Parse(global.SHUKKOU_ID))
        //        {
        //            if (kR.社員 == global.flgOff) // 社員使用の区分か検証する
        //            {
        //                setErrStatus(eKintaiKigou, iX - 1, "社員が使用しない勤怠記号です");
        //                return false;
        //            }
        //        }
        //        else if (cID == int.Parse(global.PART_ID))
        //        {
        //            if (kR.パートタイマー == global.flgOff)     // パート使用の区分か検証する
        //            {
        //                setErrStatus(eKintaiKigou, iX - 1, "パートタイマーが使用しない勤怠記号です");
        //                return false;
        //            }
        //        }
        //    }

        //    return true;
        //}


        /////------------------------------------------------------------------------------------
        ///// <summary>
        /////     時間記入チェック </summary>
        ///// <param name="obj">
        /////     勤務票明細Rowコレクション</param>
        ///// <param name="tittle">
        /////     チェック項目名称</param>
        ///// <param name="Tani">
        /////     分記入単位</param>
        ///// <param name="iX">
        /////     日付を表すインデックス</param>
        ///// <param name="stKbn">
        /////     勤怠記号の出勤怠区分</param>
        ///// <returns>
        /////     エラーなし：true, エラーあり：false</returns>
        /////------------------------------------------------------------------------------------
        //private bool errCheckTime(JSDataSet.勤務票明細Row m, string tittle, int Tani, int iX, int stKbn)
        //{
        //    // 勤怠区分マスターの出勤怠区分[1]で無記入のときNGとする
        //    if (m.開始時 == string.Empty && stKbn == global.flgOn)
        //    {
        //        setErrStatus(eSH, iX - 1, tittle + "が未入力です");
        //        return false;
        //    }

        //    if (m.開始分 == string.Empty && stKbn == global.flgOn)
        //    {
        //        setErrStatus(eSM, iX - 1, tittle + "が未入力です");
        //        return false;
        //    }

        //    if (m.終了時 == string.Empty && stKbn == global.flgOn)
        //    {
        //        setErrStatus(eEH, iX - 1, tittle + "が未入力です");
        //        return false;
        //    }

        //    if (m.終了分 == string.Empty && stKbn == global.flgOn)
        //    {
        //        setErrStatus(eEM, iX - 1, tittle + "が未入力です");
        //        return false;
        //    }

        //    // 勤怠区分マスターの出勤怠区分[0]で記入済みのときNGとする
        //    if (m.開始時 != string.Empty && stKbn == global.flgOff)
        //    {
        //        setErrStatus(eSH, iX - 1, "勤怠区分が「" + m.勤怠区分 + "」で" + tittle + "が入力されています");
        //        return false;
        //    }
             
        //    if (m.開始分 != string.Empty && stKbn == global.flgOff)
        //    {
        //        setErrStatus(eSM, iX - 1, "勤怠区分が「" + m.勤怠区分 + "」で" + tittle + "が入力されています");
        //        return false;
        //    }

        //    if (m.終了時 != string.Empty && stKbn == global.flgOff)
        //    {
        //        setErrStatus(eEH, iX - 1, "勤怠区分が「" + m.勤怠区分 + "」で" + tittle + "が入力されています");
        //        return false;
        //    }

        //    if (m.終了分 != string.Empty && stKbn == global.flgOff)
        //    {
        //        setErrStatus(eEM, iX - 1, "勤怠区分が「" + m.勤怠区分 + "」で" + tittle + "が入力されています");
        //        return false;
        //    }
            
        //    // 開始時間と終了時間
        //    string sTimeW = m.開始時.Trim() + m.開始分.Trim();
        //    string eTimeW = m.終了時.Trim() + m.終了分.Trim();

        //    if (sTimeW != string.Empty && eTimeW == string.Empty)
        //    {
        //        setErrStatus(eEH, iX - 1, tittle + "退勤時刻が未入力です");
        //        return false;
        //    }

        //    if (sTimeW == string.Empty && eTimeW != string.Empty)
        //    {
        //        setErrStatus(eSH, iX - 1, tittle + "出勤時刻が未入力です");
        //        return false;
        //    }

        //    // 記入のとき
        //    if (m.開始時 != string.Empty || m.開始分 != string.Empty ||
        //        m.終了時 != string.Empty || m.終了分 != string.Empty)
        //    {
        //        // 数字範囲、単位チェック
        //        if (!checkHourSpan(m.開始時))
        //        {
        //            setErrStatus(eSH, iX - 1, tittle + "が正しくありません");
        //            return false;
        //        }

        //        if (!checkMinSpan(m.開始分, Tani))
        //        {
        //            setErrStatus(eSM, iX - 1, tittle + "が正しくありません");
        //            return false;
        //        }

        //        if (!checkHourSpan(m.終了時))
        //        {
        //            setErrStatus(eEH, iX - 1, tittle + "が正しくありません");
        //            return false;
        //        }

        //        if (!checkMinSpan(m.終了分, Tani))
        //        {
        //            setErrStatus(eEM, iX - 1, tittle + "が正しくありません");
        //            return false;
        //        }

        //        //// 終了時刻範囲
        //        //if (Utility.StrtoInt(Utility.NulltoStr(m.終了時)) == 24 &&
        //        //    Utility.StrtoInt(Utility.NulltoStr(m.終了分)) > 0)
        //        //{
        //        //    setErrStatus(eEM, iX - 1, tittle + "終了時刻範囲を超えています（～２４：００）");
        //        //    return false;
        //        //}
        //    }

        //    return true;
        //}

        /////------------------------------------------------------------------------------------
        ///// <summary>
        /////     時間記入範囲チェック 0～23の数値 </summary>
        ///// <param name="h">
        /////     記入値</param>
        ///// <returns>
        /////     正常:true, エラー:false</returns>
        /////------------------------------------------------------------------------------------
        //private bool checkHourSpan(string h)
        //{
        //    if (!Utility.NumericCheck(h)) return false;
        //    else if (int.Parse(h) < 0 || int.Parse(h) > 23) return false;
        //    else return true;
        //}

        /////------------------------------------------------------------------------------------
        ///// <summary>
        /////     分記入範囲チェック：0～59の数値及び記入単位 </summary>
        ///// <param name="h">
        /////     記入値</param>
        ///// <param name="tani">
        /////     記入単位分</param>
        ///// <returns>
        /////     正常:true, エラー:false</returns>
        /////------------------------------------------------------------------------------------
        //private bool checkMinSpan(string m, int tani)
        //{
        //    if (!Utility.NumericCheck(m)) return false;
        //    else if (int.Parse(m) < 0 || int.Parse(m) > 59) return false;
        //    else if (int.Parse(m) % tani != 0) return false;
        //    else return true;
        //}

        /////------------------------------------------------------------------------------------
        ///// <summary>
        /////     時間外記入チェック </summary>
        ///// <param name="obj">
        /////     勤務票明細Rowコレクション</param>
        ///// <param name="tittle">
        /////     チェック項目名称</param>
        ///// <param name="Tani">
        /////     分記入単位</param>
        ///// <param name="iX">
        /////     日付を表すインデックス</param>
        ///// <returns>
        /////     エラーなし：true, エラーあり：false</returns>
        /////------------------------------------------------------------------------------------
        //private bool errCheckZan(JSDataSet.勤務票明細Row m, string tittle, int Tani, int iX)
        //{
        //    // 無記入なら終了
        //    if (m.時間外時 == string.Empty && m.時間外分 == string.Empty) return true;

        //    // 勤怠区分が出勤、有休・半日以外で記入されているときエラー
        //    if (m.時間外時 != string.Empty && m.勤怠区分 != global.K_SHUKIN && m.勤怠区分 != global.K_YUKYU_HAN)
        //    {
        //        setErrStatus(eZH, iX - 1, "出勤、有休・半日以外で" + tittle + "が入力されています");
        //        return false;
        //    }

        //    if (m.時間外分 != string.Empty && m.勤怠区分 != global.K_SHUKIN && m.勤怠区分 != global.K_YUKYU_HAN)
        //    {
        //        setErrStatus(eZM, iX - 1, "出勤、有休・半日以外で" + tittle + "が入力されています");
        //        return false;
        //    }

        //    // 出退勤時刻が無記入で時間外が記入されているときエラー
        //    if (m.開始時 == string.Empty && m.開始分 == string.Empty &&
        //         m.終了時 == string.Empty && m.終了分 == string.Empty)
        //    {
        //        if (m.時間外時 != string.Empty)
        //        {
        //            setErrStatus(eZH, iX - 1, "出退勤時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }

        //        if (m.時間外分 != string.Empty)
        //        {
        //            setErrStatus(eZM, iX - 1, "出退勤時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }
        //    }

        //    // 記入のとき
        //    if (m.時間外時 != string.Empty || m.時間外分 != string.Empty)
        //    {
        //        // 時間と分のチェック
        //        if (!checkHourSpan(m.時間外時))
        //        {
        //            setErrStatus(eZH, iX - 1, tittle + "が正しくありません");
        //            return false;
        //        }

        //        if (!checkMinSpan(m.時間外分, Tani))
        //        {
        //            setErrStatus(eZM, iX - 1, tittle + "が正しくありません。（" + Tani.ToString() + "分単位）" );
        //            return false;
        //        }
        //    }

        //    return true;
        //}

        /////------------------------------------------------------------------------------------
        ///// <summary>
        /////     休日出勤記入チェック </summary>
        ///// <param name="m">
        /////     勤務票明細Rowコレクション</param>
        ///// <param name="tittle">
        /////     チェック項目名称</param>
        ///// <param name="Tani">
        /////     分記入単位</param>
        ///// <param name="iX">
        /////     日付を表すインデックス</param>
        ///// <returns>
        /////     エラーなし：true, エラーあり：false</returns>
        /////------------------------------------------------------------------------------------
        //private bool errCheckKyujitsuShukkin(JSDataSet.勤務票明細Row m, string tittle, int Tani, int iX)
        //{
        //    // 「休日出勤・代休無し」または「休日出勤・代休取得」で無記入のときエラー
        //    if ( (m.勤怠区分 == global.K_KYUJITSUSHUKIN || m.勤怠区分 == global.K_KYUJITSUSHUKIN_D))
        //    {
        //        if (m.休日出勤時 == string.Empty )
        //        {
        //            setErrStatus(eKSH, iX - 1, "休日出勤で" + tittle + "が無記入です");
        //            return false;
        //        }
            
        //        if (m.休日出勤分 == string.Empty)
        //        {
        //            setErrStatus(eKSM, iX - 1, "休日出勤で" + tittle + "が無記入です");
        //            return false;
        //        }
        //    }

        //    // 「休日出勤・代休無し」または「休日出勤・代休取得」以外で記入済のときエラー
        //    if (m.勤怠区分 != global.K_KYUJITSUSHUKIN && m.勤怠区分 != global.K_KYUJITSUSHUKIN_D)
        //    {
        //        if (m.休日出勤時 != string.Empty)
        //        {
        //            setErrStatus(eKSH, iX - 1, "休日出勤以外で" + tittle + "が記入されています");
        //            return false;
        //        }
                
        //        if (m.休日出勤分 != string.Empty)
        //        {
        //            setErrStatus(eKSM, iX - 1, "休日出勤以外で" + tittle + "が記入されています");
        //            return false;
        //        }
        //    }

        //    // 出退勤時刻が無記入で時間外が記入されているときエラー
        //    if (m.開始時 == string.Empty && m.開始分 == string.Empty &&
        //         m.終了時 == string.Empty && m.終了分 == string.Empty)
        //    {
        //        if (m.休日出勤時 != string.Empty)
        //        {
        //            setErrStatus(eKSH, iX - 1, "出退勤時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }

        //        if (m.休日出勤分 != string.Empty)
        //        {
        //            setErrStatus(eKSM, iX - 1, "出退勤時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }
        //    }

        //    // 記入があるとき
        //    if (m.休日出勤時 != string.Empty || m.休日出勤分 != string.Empty)
        //    {
        //        // 時間と分のチェック
        //        if (!checkHourSpan(m.休日出勤時))
        //        {
        //            setErrStatus(eKSH, iX - 1, tittle + "が正しくありません");
        //            return false;
        //        }

        //        if (!checkMinSpan(m.休日出勤分, Tani))
        //        {
        //            setErrStatus(eKSM, iX - 1, tittle + "が正しくありません（" + Tani.ToString() + "分単位）");
        //            return false;
        //        }

        //        // 代休取得は６時間以上の勤務が必要
        //        if (m.勤怠区分 == global.K_KYUJITSUSHUKIN_D)
        //        {
        //            if (Utility.StrtoInt(m.休日出勤時) < 6)
        //            {
        //                setErrStatus(eKSH, iX - 1, tittle + "が６時間未満で代休取得は出来ません");
        //                return false;
        //            }
        //        }
        //    }
            
        //    return true;
        //}

        /////------------------------------------------------------------------------------------
        ///// <summary>
        /////     深夜勤務記入チェック </summary>
        ///// <param name="m">
        /////     勤務票明細Rowコレクション</param>
        ///// <param name="tittle">
        /////     チェック項目名称</param>
        ///// <param name="Tani">
        /////     分記入単位</param>
        ///// <param name="iX">
        /////     日付を表すインデックス</param>
        ///// <returns>
        /////     エラーなし：true, エラーあり：false</returns>
        /////------------------------------------------------------------------------------------
        //private bool errCheckShinya(JSDataSet.勤務票明細Row m, string tittle, int Tani, int iX)
        //{
        //    // 「出勤」「休日出勤・代休無し」「休日出勤・代休取得」「有休・半日」以外で記入済のときエラー
        //    if (m.勤怠区分 != global.K_SHUKIN && m.勤怠区分 != global.K_KYUJITSUSHUKIN && 
        //        m.勤怠区分 != global.K_KYUJITSUSHUKIN_D && m.勤怠区分 != global.K_YUKYU_HAN)
        //    {
        //        if (m.深夜時 != string.Empty)
        //        {
        //            setErrStatus(eSIH, iX - 1, "出勤、休日出勤、有休・半日以外で" + tittle + "が記入されています");
        //            return false;
        //        }

        //        if (m.深夜分 != string.Empty)
        //        {
        //            setErrStatus(eSIM, iX - 1, "出勤、休日出勤、有休・半日以外で" + tittle + "が記入されています");
        //            return false;
        //        }
        //    }

        //    // 出退勤時刻が無記入で深夜が記入されているときエラー
        //    if (m.開始時 == string.Empty && m.開始分 == string.Empty &&
        //         m.終了時 == string.Empty && m.終了分 == string.Empty)
        //    {
        //        if (m.深夜時 != string.Empty)
        //        {
        //            setErrStatus(eSIH, iX - 1, "出退勤時刻が無記入で" + tittle + "が記入されています");
        //            return false;
        //        }

        //        if (m.深夜分 != string.Empty)
        //        {
        //            setErrStatus(eSIM, iX - 1, "出退勤時刻が無記入で" + tittle + "が記入されています");
        //            return false;
        //        }
        //    }

        //    // 深夜稼働時間取得
        //    _workShinyaTime = getShinyaWorkTime(m.開始時, m.開始分, m.終了時, m.終了分);

        //    // 深夜稼働時間をTimeSpan型で取得
        //    TimeSpan siSpan = TimeSpan.FromMinutes(_workShinyaTime);

        //    // 記入があるとき
        //    if (m.深夜時 != string.Empty || m.深夜分 != string.Empty)
        //    {
        //        // 時間と分のチェック
        //        if (!checkHourSpan(m.深夜時))
        //        {
        //            setErrStatus(eSIH, iX - 1, tittle + "が正しくありません");
        //            return false;
        //        }

        //        if (!checkMinSpan(m.深夜分, Tani))
        //        {
        //            setErrStatus(eSIM, iX - 1, tittle + "が正しくありません。（" + Tani.ToString() +"分単位）");
        //            return false;
        //        }
                
        //        if (_workShinyaTime == 0 && m.深夜時 != string.Empty)
        //        {
        //            setErrStatus(eSIH, iX - 1, "深夜時間帯の勤務はありません");
        //            return false;
        //        }

        //        if (_workShinyaTime == 0 && m.深夜分 != string.Empty)
        //        {
        //            setErrStatus(eSIM, iX - 1, "深夜時間帯の勤務はありません");
        //            return false;
        //        }

        //        // 記入深夜時間
        //        double kShinya = Utility.StrtoDouble(m.深夜時) * 60 + Utility.StrtoDouble(m.深夜分);
        //        if (_workShinyaTime < kShinya)
        //        {
        //            setErrStatus(eSIH, iX - 1, "実際の深夜勤務（" + siSpan.Hours.ToString() + "時間" + siSpan.Minutes.ToString() + "分)より多く記入されています");
        //            return false;
        //        }

        //        // 深夜勤務時間はMAX７時間
        //        if (kShinya > 420)
        //        {
        //            setErrStatus(eSIH, iX - 1, tittle + "が７時間を超えています");
        //            return false;
        //        }
        //    }

        //    // 30分以上深夜勤務があり無記入のとき
        //    if (m.深夜時 == string.Empty && m.深夜分 == string.Empty && _workShinyaTime >= 30)
        //    {
        //        setErrStatus(eSIH, iX - 1, siSpan.Hours.ToString() + "時間" + siSpan.Minutes.ToString() + "分の深夜勤務がありますが無記入です");
        //        return false;
        //    }

        //    return true;
        //}

        /////------------------------------------------------------------------------------------
        ///// <summary>
        /////     実働時間を取得する</summary>
        ///// <param name="m">
        /////     勤務票明細Rowコレクション</param>
        ///// <param name="sH">
        /////     開始時</param>
        ///// <param name="sM">
        /////     開始分</param>
        ///// <param name="eH">
        /////     終了時</param>
        ///// <param name="eM">
        /////     終了分</param>
        ///// <param name="rH">
        /////     休憩時間・分</param>
        ///// <returns>
        /////     実働時間</returns>
        /////------------------------------------------------------------------------------------
        //public double getWorkTime(string sH, string sM, string eH, string eM, int rH)
        //{
        //    DateTime sTm;
        //    DateTime eTm;
        //    DateTime cTm;
        //    double w = 0;   // 稼働時間

        //    // 時刻情報に不備がある場合は０を返す
        //    if (!Utility.NumericCheck(sH) || !Utility.NumericCheck(sM) || 
        //        !Utility.NumericCheck(eH) || !Utility.NumericCheck(eM))
        //        return 0;

        //    // 開始時刻取得
        //    if (Utility.StrtoInt(sH) == 24)
        //    {
        //        if (DateTime.TryParse("0:" + Utility.StrtoInt(sM).ToString(), out cTm))
        //        {
        //            sTm = cTm;
        //        }
        //        else return 0;
        //    }
        //    else
        //    {
        //        if (DateTime.TryParse(Utility.StrtoInt(sH).ToString() + ":" + Utility.StrtoInt(sM).ToString(), out cTm))
        //        {
        //            sTm = cTm;
        //        }
        //        else return 0;
        //    }

        //    // 終了時刻取得
        //    if (Utility.StrtoInt(eH) == 24)
        //        eTm = DateTime.Parse("23:59");
        //    else
        //    {
        //        if (DateTime.TryParse(Utility.StrtoInt(eH).ToString() + ":" + Utility.StrtoInt(eM).ToString(), out cTm))
        //        {
        //            eTm = cTm;
        //        }
        //        else return 0;
        //    }

        //    // 終了時間が24:00記入のときは23:59までの計算なので稼働時間1分加算する
        //    if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
        //        w = Utility.GetTimeSpan(sTm, eTm).TotalMinutes + 1;
        //    else w = Utility.GetTimeSpan(sTm, eTm).TotalMinutes;  // 稼働時間

        //    // 休憩時間を差し引く
        //    if (w >= rH) w = w - rH;
        //    else w = 0;

        //    // 値を返す
        //    return w;
        //}

        /////--------------------------------------------------------------
        ///// <summary>
        /////     深夜勤務時間を取得する</summary>
        ///// <param name="m">
        /////     勤務票明細Rowコレクション</param>
        ///// <param name="sH">
        /////     開始時</param>
        ///// <param name="sM">
        /////     開始分</param>
        ///// <param name="eH">
        /////     終了時</param>
        ///// <param name="eM">
        /////     終了分</param>
        ///// <returns>
        /////     深夜勤務時間</returns>
        ///// ------------------------------------------------------------
        //private double getShinyaWorkTime(string sH, string sM, string eH, string eM)
        //{
        //    DateTime sTime;
        //    DateTime eTime;
        //    DateTime cTm;

        //    double wkShinya = 0;    // 深夜稼働時間

        //    // 時刻情報に不備がある場合は０を返す
        //    if (!Utility.NumericCheck(sH) || !Utility.NumericCheck(sM) ||
        //        !Utility.NumericCheck(eH) || !Utility.NumericCheck(eM))
        //        return 0;

        //    // 開始時間を取得
        //    if (DateTime.TryParse(Utility.StrtoInt(sH).ToString() + ":" + Utility.StrtoInt(sM).ToString(), out cTm))
        //    {
        //        sTime = cTm;
        //    }
        //    else return 0;

        //    // 終了時間を取得
        //    if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
        //    {
        //        eTime = global.dt2359;
        //    }
        //    else if (DateTime.TryParse(Utility.StrtoInt(eH).ToString() + ":" + Utility.StrtoInt(eM).ToString(), out cTm))
        //    {
        //        eTime = cTm;
        //    }
        //    else return 0;


        //    // 当日内の勤務のとき
        //    if (sTime.TimeOfDay <= eTime.TimeOfDay)
        //    {
        //        // 早出残業時間を求める
        //        if (sTime < global.dt0500)  // 開始時刻が午前5時前のとき
        //        {
        //            // 早朝時間帯稼働時間
        //            if (eTime >= global.dt0500)
        //            {
        //                wkShinya += Utility.GetTimeSpan(sTime, global.dt0500).TotalMinutes;
        //            }
        //            else
        //            {
        //                wkShinya += Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
        //            }
        //        }

        //        // 終了時刻が22:00以降のとき
        //        if (eTime >= global.dt2200)
        //        {
        //            // 当日分の深夜帯稼働時間を求める
        //            if (sTime <= global.dt2200)
        //            {
        //                // 出勤時刻が22:00以前のとき深夜開始時刻は22:00とする
        //                wkShinya += Utility.GetTimeSpan(global.dt2200, eTime).TotalMinutes;
        //            }
        //            else
        //            {
        //                // 出勤時刻が22:00以降のとき深夜開始時刻は出勤時刻とする
        //                wkShinya += Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
        //            }

        //            // 終了時間が24:00記入のときは23:59までの計算なので稼働時間1分加算する
        //            if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
        //                wkShinya += 1;
        //        }
        //    }
        //    else
        //    {
        //        // 日付を超えて終了したとき（開始時刻 > 終了時刻）

        //        // 早出残業時間を求める
        //        if (sTime < global.dt0500)  // 開始時刻が午前5時前のとき
        //        {
        //            wkShinya += Utility.GetTimeSpan(sTime, global.dt0500).TotalMinutes;
        //        }

        //        // 当日分の深夜勤務時間（～０：００まで）
        //        if (sTime <= global.dt2200)
        //        {
        //            // 出勤時刻が22:00以前のとき無条件に120分
        //            wkShinya += global.TOUJITSU_SINYATIME;
        //        }
        //        else
        //        {
        //            // 出勤時刻が22:00以降のとき出勤時刻から24:00までを求める
        //            wkShinya += Utility.GetTimeSpan(sTime, global.dt2359).TotalMinutes + 1;
        //        }

        //        // 0:00以降の深夜勤務時間を加算（０：００～終了時刻）
        //        if (eTime.TimeOfDay > global.dt0500.TimeOfDay)
        //        {
        //            wkShinya += Utility.GetTimeSpan(global.dt0000, global.dt0500).TotalMinutes;
        //        }
        //        else
        //        {
        //            wkShinya += Utility.GetTimeSpan(global.dt0000, eTime).TotalMinutes;
        //        }
        //    }

        //    return wkShinya;
        //}

        /////----------------------------------------------------------------------
        ///// <summary>
        /////     呼出コードチェック</summary>
        ///// <param name="m">
        /////     勤務票明細Rowコレクション</param>
        ///// <param name="tittle">
        /////     チェック項目名称</param>
        ///// <param name="iX">
        /////     日付を表すインデックス</param>
        ///// <returns>
        /////     エラーなし：true, エラーあり：false</returns>
        /////----------------------------------------------------------------------
        //private bool errCheckYCD(JSDataSet.勤務票明細Row m, string tittle, int iX)
        //{
        //    if (m.呼出 == 0) return true;

        //    // 出勤、休日出勤以外で記入されているとエラー
        //    if (m.勤怠区分 != global.K_SHUKIN && m.勤怠区分 != global.K_KYUJITSUSHUKIN &&
        //        m.勤怠区分 != global.K_KYUJITSUSHUKIN_D)
        //    {
        //        setErrStatus(eYCD, iX - 1, "出勤、休日出勤以外で" + tittle + "が記入されています");
        //        return false;
        //    }

        //    // 存在しない区分のときエラー
        //    if (m.呼出 != global.YOBICODE_1 && m.呼出 != global.YOBICODE_2)
        //    {
        //        setErrStatus(eYCD, iX - 1, tittle + "が正しくありません");
        //        return false;
        //    }

        //    return true;
        //}

        /////----------------------------------------------------------------------
        ///// <summary>
        /////     交替コードチェック</summary>
        ///// <param name="m">
        /////     勤務票明細Rowコレクション</param>
        ///// <param name="tittle">
        /////     チェック項目名称</param>
        ///// <param name="iX">
        /////     日付を表すインデックス</param>
        ///// <returns>
        /////     エラーなし：true, エラーあり：false</returns>
        /////---------------------------------------------------------------------- 
        //private bool errCheckKotaiCD(JSDataSet.勤務票明細Row m, string tittle, int iX)
        //{
        //    if (m.交替 == 0) return true;

        //    if (m.交替 != global.KOUTAI_ASA && m.交替 != global.KOUTAI_NAKA && m.交替  != global.KOUTAI_YORU)
        //    {
        //        setErrStatus(eKCD, iX - 1, tittle + "が正しくありません");
        //        return false;
        //    }

        //    // 「出勤」「休日出勤・代休無し」「休日出勤・代休取得」「有休・半日」以外で記入されているとエラー
        //    if (m.交替 != 0 && 
        //        m.勤怠区分 != global.K_SHUKIN && m.勤怠区分 != global.K_KYUJITSUSHUKIN &&
        //        m.勤怠区分 != global.K_KYUJITSUSHUKIN_D && m.勤怠区分 != global.K_YUKYU_HAN)
        //    {
        //        setErrStatus(eKCD, iX - 1, "出勤、休日出勤、有休・半日以外で" + tittle + "が記入されています");
        //        return false;
        //    }

        //    // 有休・半日で「夜番」はエラー
        //    if (m.勤怠区分 == global.K_YUKYU_HAN && m.交替 == global.KOUTAI_YORU)
        //    {
        //        setErrStatus(eKCD, iX - 1, "有休・半日で「夜番」は記入できません");
        //        return false;
        //    }

        //    return true;
        //}

        /////--------------------------------------------------------------------------- 
        ///// <summary>
        /////     遅刻早退外出記入チェック</summary>
        ///// <param name="obj">
        /////     勤務票明細Rowコレクション</param>
        ///// <param name="tittle">
        /////     チェック項目名称</param>
        ///// <param name="Tani">
        /////     分記入単位</param>
        ///// <param name="iX">
        /////     日付を表すインデックス</param>
        ///// <returns>
        /////     エラーなし：true, エラーあり：false</returns>
        /////--------------------------------------------------------------------------- 
        //private bool errCheckChikoku(JSDataSet.勤務票明細Row m, string tittle, int Tani, int iX)
        //{
        //    // 無記入なら終了
        //    if (m.遅早退時 == string.Empty && m.遅早退分 == string.Empty) return true;

        //    // 勤怠区分が「出勤」「有休・半日」以外で記入されているときエラー
        //    if (m.勤怠区分 != global.K_SHUKIN && m.勤怠区分 != global.K_YUKYU_HAN)
        //    {
        //        if (m.遅早退時 != string.Empty)
        //        {
        //            setErrStatus(eCSGH, iX - 1, "出勤、有休・半日以外で" + tittle + "が入力されています");
        //            return false;
        //        }

        //        if (m.遅早退分 != string.Empty)
        //        {
        //            setErrStatus(eCSGM, iX - 1, "出勤、有休・半日以外で" + tittle + "が入力されています");
        //            return false;
        //        }
        //    }
            
        //    // 出退勤時刻が無記入で記入されているとエラー
        //    if (m.開始時 == string.Empty && m.開始分 == string.Empty &&
        //         m.終了時 == string.Empty && m.終了分 == string.Empty)
        //    {
        //        if (m.遅早退時 != string.Empty)
        //        {
        //            setErrStatus(eCSGH, iX - 1, "出退勤時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }

        //        if (m.遅早退分 != string.Empty)
        //        {
        //            setErrStatus(eCSGM, iX - 1, "出退勤時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }
        //    }

        //    // 記入があるとき
        //    if (m.遅早退時 != string.Empty || m.遅早退分 != string.Empty)
        //    {
        //        // 時間と分のチェック
        //        if (!checkHourSpan(m.遅早退時))
        //        {
        //            setErrStatus(eCSGH, iX - 1, tittle + "が正しくありません");
        //            return false;
        //        }

        //        if (!checkMinSpan(m.遅早退分, Tani))
        //        {
        //            setErrStatus(eCSGM, iX - 1, tittle + "が正しくありません。（" + Tani.ToString() + "分単位）");
        //            return false;
        //        }
        //    }

        //    return true;
        //}

        /////--------------------------------------------------------------------------- 
        ///// <summary>
        /////     カット時間記入チェック</summary>
        ///// <param name="obj">
        /////     勤務票ヘッダRowコレクション</param>
        ///// <param name="tittle">
        /////     チェック項目名称</param>
        ///// <param name="Tani">
        /////     分記入単位</param>
        ///// <returns>
        /////     エラーなし：true, エラーあり：false</returns>
        /////--------------------------------------------------------------------------- 
        //private bool errCheckCutTime(JSDataSet.勤務票ヘッダRow r, string tittle, int Tani)
        //{
        //    // 無記入なら終了
        //    if (r.カット時 == 0 && r.カット分 == 0) return true;

        //    // 時間チェック
        //    if (!Utility.NumericCheck(r.カット時.ToString()))
        //    {
        //        setErrStatus(eCUTH, 0, tittle + "が正しくありません。");
        //    }

        //    // 分のチェック
        //    if (!checkMinSpan(r.カット分.ToString(), Tani))
        //    {
        //        setErrStatus(eCUTM, 0, tittle + "が正しくありません。（" + Tani.ToString() + "分単位）");
        //        return false;
        //    }

        //    return true;
        //}

        /////-----------------------------------------------------------------------------
        ///// <summary>
        /////     特殊作業時間記入チェック</summary>
        ///// <param name="obj">
        /////     勤務票ヘッダRowコレクション</param>
        ///// <param name="tittle">
        /////     チェック項目名称</param>
        ///// <param name="Tani">
        /////     分記入単位</param>
        ///// <returns>
        /////     エラーなし：true, エラーあり：false</returns>
        /////-----------------------------------------------------------------------------
        //private bool errCheckTokushu(JSDataSet.勤務票ヘッダRow r, string tittle, int Tani)
        //{
        //    // 無記入なら終了
        //    if (r.特殊手当時 == 0 && r.特殊手当分 == 0) return true;

        //    // 時間と分のチェック
        //    if (!Utility.NumericCheck(r.特殊手当時.ToString()))
        //    {
        //        setErrStatus(eTKSH, 0, tittle + "が正しくありません");
        //        return false;
        //    }

        //    // 分のチェック
        //    if (!checkMinSpan(r.特殊手当分.ToString(), Tani))
        //    {
        //        setErrStatus(eTKSM, 0, tittle + "が正しくありません。（" + Tani.ToString() + "分単位）");
        //        return false;
        //    }

        //    return true;
        //}

        /////-----------------------------------------------------------------------------
        ///// <summary>
        /////     所定内休日記入チェック（パート）</summary>
        ///// <param name="obj">
        /////     勤務票明細Rowコレクション</param>
        ///// <param name="tittle">
        /////     チェック項目名称</param>
        ///// <param name="Tani">
        /////     分記入単位</param>
        ///// <param name="iX">
        /////     日付を表すインデックス</param>
        ///// <returns>
        /////     エラーなし：true, エラーあり：false</returns>
        /////-----------------------------------------------------------------------------
        //private bool errCheckShotei(JSDataSet.勤務票明細Row m, string tittle, int Tani, int iX)
        //{
        //    // 勤怠区分が「出勤」「休日出勤・代休無し」「有休・半日」以外で記入されているときエラー
        //    if (m.勤怠区分 != global.K_SHUKIN && m.勤怠区分 != global.K_KYUJITSUSHUKIN && m.勤怠区分 != global.K_YUKYU_HAN)
        //    {
        //        if (m.所定内休日時 != string.Empty)
        //        {
        //            setErrStatus(eSHOH, iX - 1, "出勤、休日出勤、有休・半日以外で" + tittle + "が入力されています");
        //            return false;
        //        }

        //        if (m.所定内休日分 != string.Empty)
        //        {
        //            setErrStatus(eSHOM, iX - 1, "出勤、休日出勤、有休・半日以外で" + tittle + "が入力されています");
        //            return false;
        //        }
        //    }

        //    // 出退勤時刻が無記入で記入されているとエラー
        //    if (m.開始時 == string.Empty && m.開始分 == string.Empty &&
        //         m.終了時 == string.Empty && m.終了分 == string.Empty)
        //    {
        //        if (m.所定内休日時 != string.Empty)
        //        {
        //            setErrStatus(eSHOH, iX - 1, "出退勤時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }

        //        if (m.所定内休日分 != string.Empty)
        //        {
        //            setErrStatus(eSHOM, iX - 1, "出退勤時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }
        //    }

        //    // 記入のとき
        //    if (m.所定内休日時 != string.Empty || m.所定内休日分 != string.Empty)
        //    {
        //        // 時間と分のチェック
        //        if (!checkHourSpan(m.所定内休日時))
        //        {
        //            setErrStatus(eSHOH, iX - 1, tittle + "が正しくありません");
        //            return false;
        //        }

        //        if (!checkMinSpan(m.所定内休日分, Tani))
        //        {
        //            setErrStatus(eSHOM, iX - 1, tittle + "が正しくありません。（" + Tani.ToString() + "分単位）");
        //            return false;
        //        }
        //    }

        //    return true;
        //}
    }
}
