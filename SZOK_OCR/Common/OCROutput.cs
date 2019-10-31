using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace JS_OCR.Common
{
    ///------------------------------------------------------------------
    /// <summary>
    ///     給与大臣受け渡しデータクラス </summary>
    ///     
    ///------------------------------------------------------------------
    class OCROutput
    {
        #region フィールド定義
        // 親フォーム
        Form _preForm;

        // 休日出勤の7.5時間勤務（分）
        const int work75 = 450;
        #endregion

        #region データテーブルインスタンス
        JSDataSet.勤務票ヘッダDataTable _hTbl;
        JSDataSet.勤務票明細DataTable _mTbl;
        #endregion

        #region 受け渡しデータファイル名
        private const string TXTFILE_SHAIN = "社員勤怠";    // 社員
        private const string TXTFILE_PART = "パート勤怠";   // パート
        private const string TXTFILE_SHUKKO = "出向社員";   // 出向社員
        #endregion

        #region 社員・出向社員　給与大臣受入データプロパティ
        //---------------------------------------------------
        //      給与大臣受入データプロパティ：社員・出向社員
        //---------------------------------------------------
        /// <summary>汎用データ：社員番号</summary>
        public string sShainNo { get; private set; }

        /// <summary>休日出勤時間・分</summary>
        public int sKyuShutsuTimes { get; private set; }

        /// <summary>休日出勤・代休あり日数</summary>
        public int sKyuShutsuDays_d { get; private set; }

        /// <summary>休日出勤・代休なし日数</summary>
        public int sKyuShutsuDays { get; private set; }

        /// <summary>使用しない</summary>
        public string dummy1 { get; private set; }

        /// <summary>ストック休暇</summary>
        public int sStockKyuuka { get; private set; }

        /// <summary>特殊作業時間・分</summary>
        public int sTokushuTeate { get; private set; }

        /// <summary>特別休暇・無給</summary>
        public int sTokukhuDays_Non { get; private set; }

        /// <summary>休職日数</summary>
        public int sKyushokuDays { get; private set; }

        /// <summary>出勤日数</summary>
        public int sShukkinDays { get; private set; }

        /// <summary>休日日数</summary>
        public int sKyujitsuDays { get; private set; }

        /// <summary>半休日数</summary>
        public int sHanKyuDays { get; private set; }

        /// <summary>有休日数</summary>
        public int sYuKyuDays { get; private set; }

        /// <summary>代休日数</summary>
        public int sDaiKyuDays { get; private set; }

        /// <summary>特別休暇日数</summary>
        public int sTokuKyuDays { get; private set; }

        /// <summary>公傷日数</summary>
        public int sKoshoDays { get; private set; }

        /// <summary>出張日数</summary>
        public int sShuchoDays { get; private set; }

        /// <summary>欠勤日数</summary>
        public int sKekkinDays { get; private set; }

        /// <summary>早出残業時間・分</summary>
        public int sZanTimes { get; private set; }

        /// <summary>使用しない</summary>
        public string dummy2 { get; private set; }

        /// <summary>代休時間・分</summary>
        public int sDaiKyuTimes { get; private set; }

        /// <summary>深夜時間・分</summary>
        public int sShinyaTimes { get; private set; }

        /// <summary>呼出・昼回数</summary>
        public int sYobi1 { get; private set; }

        /// <summary>呼出・夜回数</summary>
        public int sYobi2 { get; private set; }

        /// <summary>朝番回数</summary>
        public int sAsaCount { get; private set; }

        /// <summary>中番回数</summary>
        public int sNakaCount { get; private set; }

        /// <summary>夜番回数</summary>
        public int sYoruCount { get; private set; }

        /// <summary>遅早時間・分</summary>
        public int sChisouTimes { get; private set; }

        /// <summary>カット時間・分</summary>
        public int sCutTimes { get; private set; }

        #endregion

        #region パートタイマー給与大臣受入データプロパティ

        //---------------------------------------------------
        //      給与大臣受入データプロパティ：パート
        //---------------------------------------------------
        /// <summary>汎用データ：社員番号</summary>
        public string pShainNo { get; private set; }

        /// <summary>出勤日数</summary>
        public int pShukkinDays { get; private set; }

        /// <summary>欠勤日数</summary>
        public int pKekkinDays { get; private set; }

        /// <summary>指定休日数</summary>
        public int pShiteiKyuDays { get; private set; }

        /// <summary>有休日数</summary>
        public int pYuKyuDays { get; private set; }

        /// <summary>公暇日数</summary>
        public int pkoukaDays { get; private set; }

        /// <summary>労働時間・分：「出勤」の日の「所定内・休日」合計</summary>
        public int pShoteiTimes { get; private set; }

        /// <summary>残業1.2・分</summary>
        public int pZan12Times { get; private set; }

        /// <summary>残業1.25・分</summary>
        public int pZan125Times { get; private set; }

        /// <summary>特別手当・分</summary>
        public int pTokubetsuTeate { get; private set; }

        /// <summary>深夜時間・分</summary>
        public int pShinyaTimes { get; private set; }

        /// <summary>休日出勤時間・分：「休日出勤・代休無し」日の「所定内・休日」合計</summary>
        public int pKyuShutsuTimes { get; private set; }

        /// <summary>半休日数</summary>
        public int pHanKyuDays { get; private set; }

        /// <summary>公傷日数</summary>
        public int pKoshoDays { get; private set; }

        /// <summary>特別手当区分</summary>
        public int pTokubetsuKbn { get; private set; }

        /// <summary>労働時間・半休：「有休・半日」日の「所定内・休日」合計・分</summary>
        public int pShoteiHanTimes { get; private set; }

        // 2014/05/07 追加
        /// <summary>休日出勤日数：「休日出勤・代休無し」</summary>
        public int pKyuShutsuDays { get; private set; }

       #endregion

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     給与大臣受入データ作成クラスコンストラクタ</summary>
        /// <param name="preFrm">
        ///     親フォーム</param>
        /// <param name="hTbl">
        ///     勤務票ヘッダDataTable</param>
        /// <param name="mTbl">
        ///     勤務票明細DataTable</param>
        ///--------------------------------------------------------------------------
        public OCROutput(Form preFrm, JSDataSet.勤務票ヘッダDataTable hTbl, JSDataSet.勤務票明細DataTable mTbl)
        {
            _preForm = preFrm;
            _hTbl = hTbl;
            _mTbl = mTbl;
        }

        ///--------------------------------------------------------------------------------------
        /// <summary>
        ///     給与大臣受入データ作成</summary>
        ///--------------------------------------------------------------------------------------     
        public void SaveData()
        {
            #region 出力配列
            string[] arrayShain = null;     // 社員出力配列
            string[] arrayPart = null;      // パート出力配列
            string[] arrayShukko = null;    // 出向社員出力配列
            #endregion

            #region 出力件数変数
            int sCnt = 0;   // 社員出力件数
            int pCnt = 0;   // パート出力件数
            int uCnt = 0;   // 出向社員出力件数
            #endregion

            Boolean pblFirstGyouFlg = true;
            string wID = string.Empty;

            // 出力先フォルダがあるか？なければ作成する
            string cPath = global.cnfPath;
            if (!System.IO.Directory.Exists(cPath)) System.IO.Directory.CreateDirectory(cPath);

            try
            {
                //オーナーフォームを無効にする
                _preForm.Enabled = false;

                //プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = _preForm;
                frmP.Show();

                // 勤務票ヘッダレコード件数取得
                int cTotal = _hTbl.Count();
                int rCnt = 1;

                // 伝票最初行フラグ
                pblFirstGyouFlg = true;

                // 勤務票ヘッダデータ取得
                var s = _hTbl.OrderBy(a => a.ID);

                foreach (var r in s)
                {
                    // プログレスバー表示
                    frmP.Text = "給与大臣受入データ作成中です・・・" + rCnt.ToString() + "/" + cTotal.ToString();
                    frmP.progressValue = rCnt * 100 / cTotal;
                    frmP.ProgressStep();

                    // 帳票番号別の準備
                    switch (r.帳票番号.ToString())
                    {
                        case global.SHAIN_ID:   // 社員

                            // 出力データ初期化
                            InitOutRecShainShukko();

                            // 社員番号
                            sShainNo = r.個人番号.ToString().PadLeft(global.ShainMaxLength, '0');

                            // カット時間
                            sCutTimes = r.カット時 * 60 + r.カット分;

                            // 特殊手当時間
                            sTokushuTeate = r.特殊手当時 * 60 + r.特殊手当分;

                            break;

                        case global.PART_ID:    // パート

                            // 出力データ初期化
                            InitOutRecPart();

                            // 社員番号
                            pShainNo = r.個人番号.ToString().PadLeft(global.ShainMaxLength, '0');

                            // 特別手当区分
                            pTokubetsuKbn = r.特別手当区分;

                            break;

                        case global.SHUKKOU_ID:     // 出向社員

                            // 出力データ初期化
                            InitOutRecShainShukko();

                            // 社員番号
                            sShainNo = r.個人番号.ToString().PadLeft(global.ShainMaxLength, '0');

                            // カット時間
                            sCutTimes = r.カット時 * 60 + r.カット分;

                            // 特殊手当時間
                            sTokushuTeate = r.特殊手当時 * 60 + r.特殊手当分;

                            break;

                        default:
                            break;
                    }

                    // 勤務票明細データ取得
                    var m = _mTbl.Where(a => a.ヘッダID == r.ID);

                    // 明細集計処理
                    foreach (var t in m)
                    {
                        switch (r.帳票番号.ToString())
                        {
                            case global.SHAIN_ID:   // 社員
                                SetDataShainShukko(t);
                                break;

                            case global.PART_ID:    // パート
                                SetDataPart(t);
                                break;

                            case global.SHUKKOU_ID: // 出向社員
                                SetDataShainShukko(t);
                                break;

                            default:
                                break;
                        }
                    }

                    // 配列にデータを出力
                    switch (r.帳票番号.ToString())
                    {
                        case global.SHAIN_ID:   // 社員
                            sCnt++;
                            Array.Resize(ref arrayShain, sCnt);
                            arrayShain[sCnt - 1] = getTextShainData();
                            break;

                        case global.PART_ID:    // パート
                            pCnt++;
                            Array.Resize(ref arrayPart, pCnt);
                            arrayPart[pCnt - 1] = getTextPartData();
                            break;

                        case global.SHUKKOU_ID: // 出向社員                            
                            uCnt++;
                            Array.Resize(ref arrayShukko, uCnt);
                            arrayShukko[uCnt - 1] = getTextShukkoData();
                            break;

                        default:
                            break;
                    }

                    //データ件数加算
                    rCnt++;

                    pblFirstGyouFlg = false;
                }

                // 社員勤怠テキストファイル出力
                if (arrayShain != null) txtFileWrite(cPath + TXTFILE_SHAIN + ".txt", arrayShain);

                // パート勤怠テキストファイル出力
                if (arrayPart != null) txtFileWrite(cPath + TXTFILE_PART + ".txt", arrayPart);

                // 出向社員勤怠テキストファイル出力
                if (arrayShukko != null) txtFileWrite(cPath + TXTFILE_SHUKKO + ".txt", arrayShukko);

                // いったんオーナーをアクティブにする
                _preForm.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                _preForm.Enabled = true;

            }
            catch (Exception e)
            {
                MessageBox.Show("給与大臣受入データ作成中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                //if (OutData.sCom.Connection.State == ConnectionState.Open) OutData.sCom.Connection.Close();
            }
        }


        ///----------------------------------------------------------------------------
        /// <summary>
        ///     配列にテキストデータをセットする </summary>
        /// <param name="array">
        ///     社員、パート、出向社員の各配列</param>
        /// <param name="cnt">
        ///     拡張する配列サイズ</param>
        /// <param name="txtData">
        ///     セットする文字列</param>
        ///----------------------------------------------------------------------------
        private void txtArraySet(string [] array, int cnt, string txtData)
        {
            Array.Resize(ref array, cnt);   // 配列のサイズ拡張
            array[cnt - 1] = txtData;       // 文字列のセット
        }


        ///----------------------------------------------------------------------------
        /// <summary>
        ///     テキストファイルを出力する</summary>
        /// <param name="outFilePath">
        ///     パスを含む出力ファイル名</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        ///----------------------------------------------------------------------------
        private void txtFileWrite(string outFilePath, string [] arrayData)
        {
            // 出力ファイルが存在するとき
            if (System.IO.File.Exists(outFilePath))
            {
                // リネーム付加文字列（タイムスタンプ）
                string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
                                     DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
                                     DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0'); 

                // リネーム後ファイル名
                string reFileName = Path.GetDirectoryName(outFilePath) + @"\" + Path.GetFileNameWithoutExtension(outFilePath) + newFileName + ".txt";

                // 確認表示
                MessageBox.Show(outFilePath + "は既に存在しています。" + Environment.NewLine + Environment.NewLine + 
                    "登録済みファイルは名前を以下のように変更して保存します。" + Environment.NewLine + Environment.NewLine +
                "現）" + outFilePath + Environment.NewLine + 
                "新）" + reFileName, "既存ファイルの名前変更", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 既存のファイルをリネーム
                File.Move(outFilePath, reFileName);
            }

            // テキストファイル出力
            File.WriteAllLines(outFilePath, arrayData, System.Text.Encoding.GetEncoding(932));
        }
        
        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     社員・出向社員 給与大臣受入データ集計処理</summary>
        /// <param name="m">
        ///     JSDataSet.勤務票明細Row</param>
        ///---------------------------------------------------------------------------------- 
        private void SetDataShainShukko(JSDataSet.勤務票明細Row m)
        {
            try
            {
                switch (m.勤怠区分)
                {
                    case global.K_KYUJITSUSHUKIN:   // 休日出勤時間（分）集計
                        sKyuShutsuTimes += Utility.StrtoInt(m.休日出勤時) * 60 + Utility.StrtoInt(m.休日出勤分);

                        sKyuShutsuDays++;   //　日数
                        break;

                    case global.K_KYUJITSUSHUKIN_D: // 休日出勤時間（分）集計 ※休日出勤・代休取得のときは7.5時間超の時間数を集計                                   
                        int val = Utility.StrtoInt(m.休日出勤時) * 60 + Utility.StrtoInt(m.休日出勤分);

                        // 7時間30分超の勤務
                        if (val > work75)
                        {
                            // 休出時間（7時間30分超の時間）
                            sKyuShutsuTimes += val - work75;

                            // 代休時間（7時間30分までの時間）
                            sDaiKyuTimes += work75;
                        }
                        else
                        {
                            // 代休時間（最大7時間30分までの時間）
                            sDaiKyuTimes += val;
                        }

                        sKyuShutsuDays_d++;  //　日数
                        break;

                    case global.K_STOCK_KYUKA:  // ストック休暇日数
                        sStockKyuuka++;
                        break;

                    case global.K_TOKUBETSU_KYUKA_MUKYU:    // 特別休暇・無給日数
                        sTokukhuDays_Non++;
                        break;

                    case global.K_KYUSHOKU:     // 休職日数
                        sKyushokuDays++;
                        break;

                    case global.K_SHUKIN:       // 出勤日数
                        sShukkinDays++;
                        break;

                    case "":                    // 休日日数
                        sKyujitsuDays++;
                        break;

                    case global.K_YUKYU_HAN:    // 半休日数
                        sHanKyuDays++;
                        break;

                    case global.K_YUKYU:        // 有給日数
                        sYuKyuDays++;
                        break;

                    case global.K_DAIKYU:       // 代休日数
                        sDaiKyuDays++;
                        break;

                    case global.K_TOKUBETSU_KYUKA:  // 特別休暇日数
                        sTokuKyuDays++;
                        break;

                    case global.K_KOUSHO:           // 公傷日数
                        sKoshoDays++;
                        break;

                    case global.K_SHUCCHOU:         // 出張日数
                        sShuchoDays++;
                        break;

                    case global.K_KEKKIN:           // 欠勤日数
                        sKekkinDays++;
                        break;

                    default:
                        break;
                }

                // 早出残業時間・時間外
                sZanTimes += Utility.StrtoInt(m.時間外時) * 60 + Utility.StrtoInt(m.時間外分);

                // 深夜時間
                sShinyaTimes += Utility.StrtoInt(m.深夜時) * 60 + Utility.StrtoInt(m.深夜分);

                // 呼出区分
                if (m.呼出 == global.YOBICODE_1) sYobi1++;
                else if (m.呼出 == global.YOBICODE_2) sYobi2++;

                // 交替区分
                if (m.交替 == global.KOUTAI_ASA) sAsaCount++;
                else if (m.交替 == global.KOUTAI_NAKA) sNakaCount++;
                else if (m.交替 == global.KOUTAI_YORU) sYoruCount++;

                // 遅早時間
                sChisouTimes += Utility.StrtoInt(m.遅早退時) * 60 + Utility.StrtoInt(m.遅早退分);
            }
            catch (Exception e)
            {
                MessageBox.Show("給与大臣 社員・出向社員受入データ作成中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     パート 給与大臣受入データ集計処理</summary>
        /// <param name="m">
        ///     JSDataSet.勤務票明細Row</param>
        /// 
        ///     2014/05/07 休日出勤日数カウント機能追加   
        ///---------------------------------------------------------------------------------- 
        private void SetDataPart(JSDataSet.勤務票明細Row m)
        {
            try
            {
                switch (m.勤怠区分)
                {
                    case global.K_SHUKIN:   // 出勤日数
                        pShukkinDays++;
                        pShoteiTimes += Utility.StrtoInt(m.所定内休日時) * 60 + Utility.StrtoInt(m.所定内休日分);
                        break;

                    case global.K_KEKKIN:   // 欠勤日数
                        pKekkinDays++;
                        break;

                    case global.K_SHITEI_KYUJITSU:  // 指定休日
                        pShiteiKyuDays++;
                        break;

                    case global.K_YUKYU:    // 有給日数
                        pYuKyuDays++;
                        break;

                    case global.K_KOUKA:    // 公暇日数
                        pkoukaDays++;
                        break;

                    case global.K_KYUJITSUSHUKIN:   // 休日出勤時間（分）集計　および　日数カウント(2014/05/07追加）
                        pKyuShutsuTimes += Utility.StrtoInt(m.所定内休日時) * 60 + Utility.StrtoInt(m.所定内休日分);
                        pKyuShutsuDays++;   // 2014/05/07 追加
                        break;

                    case global.K_YUKYU_HAN:    // 半休日数
                        pHanKyuDays++;
                        pShoteiHanTimes += Utility.StrtoInt(m.所定内休日時) * 60 + Utility.StrtoInt(m.所定内休日分);
                        break;

                    case global.K_KOUSHO:   // 公傷日数
                        pKoshoDays++;
                        break;

                    default:
                        break;
                }

                // 残業1.2時間・時間外
                pZan12Times += Utility.StrtoInt(m.時間外12時) * 60 + Utility.StrtoInt(m.時間外12分);

                // 残業1.25時間・時間外
                pZan125Times += Utility.StrtoInt(m.時間外125時) * 60 + Utility.StrtoInt(m.時間外125分);

                // 特別手当時間
                pTokubetsuTeate += Utility.StrtoInt(m.特別手当時) * 60 + Utility.StrtoInt(m.特別手当分);

                // 深夜時間
                pShinyaTimes += Utility.StrtoInt(m.深夜時) * 60 + Utility.StrtoInt(m.深夜分);
            }
            catch (Exception e)
            {
                MessageBox.Show("給与大臣 パート受入データ作成中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        //-------------------------------------------------------------------------
        /// <summary>
        ///     社員勤怠.TXT レコード文字列作成</summary>
        /// <returns>
        ///     社員勤怠.TXT レコード文字列</returns>
        ///------------------------------------------------------------------------
        private string getTextShainData()
        {
            return getShainShukkoText();
        }

        //-------------------------------------------------------------------------
        /// <summary>
        ///     出向社員.TXT レコード文字列作成</summary>
        /// <returns>
        ///     出向社員.TXT レコード文字列</returns>
        ///------------------------------------------------------------------------
        private string getTextShukkoData()
        {
            return getShainShukkoText();
        }

        //-------------------------------------------------------------------------
        /// <summary>
        ///     社員勤怠.TXT・出向社員.TXT レコード文字列作成</summary>
        /// <returns>
        ///     社員勤怠.TXT, 出向社員.TXT レコード文字列</returns>
        ///------------------------------------------------------------------------
        private string getShainShukkoText()
        {
            //出力文字列作成
            StringBuilder sb = new StringBuilder();
            sb.Append(sShainNo).Append(",");
            sb.Append(sKyuShutsuTimes.ToString()).Append(",");
            sb.Append(sKyuShutsuDays_d.ToString()).Append(",");
            sb.Append(sKyuShutsuDays.ToString()).Append(",");
            sb.Append(dummy1).Append(",");
            sb.Append(sStockKyuuka.ToString()).Append(",");
            sb.Append(sTokushuTeate.ToString()).Append(",");
            sb.Append(sTokukhuDays_Non.ToString()).Append(",");
            sb.Append(sKyushokuDays.ToString()).Append(",");
            sb.Append(sShukkinDays.ToString()).Append(",");
            sb.Append(sKyujitsuDays.ToString()).Append(",");
            sb.Append(sHanKyuDays.ToString()).Append(",");
            sb.Append(sYuKyuDays.ToString()).Append(",");
            sb.Append(sDaiKyuDays.ToString()).Append(",");
            sb.Append(sTokuKyuDays.ToString()).Append(",");
            sb.Append(sKoshoDays.ToString()).Append(",");
            sb.Append(sShuchoDays.ToString()).Append(",");
            sb.Append(sKekkinDays.ToString()).Append(",");
            sb.Append(sZanTimes.ToString()).Append(",");
            sb.Append(dummy2).Append(",");
            sb.Append(sDaiKyuTimes.ToString()).Append(",");
            sb.Append(sShinyaTimes.ToString()).Append(",");
            sb.Append(sYobi1.ToString()).Append(",");
            sb.Append(sYobi2.ToString()).Append(",");
            sb.Append(sAsaCount.ToString()).Append(",");
            sb.Append(sNakaCount.ToString()).Append(",");
            sb.Append(sYoruCount.ToString()).Append(",");
            sb.Append(sChisouTimes.ToString()).Append(",");
            sb.Append(sCutTimes.ToString());

            return sb.ToString();
        }

        //-------------------------------------------------------------------------
        /// <summary>
        ///     パート勤怠.TXT レコード文字列作成</summary>
        /// <returns>
        ///     パート勤怠.TXT レコード文字列</returns>
        ///     
        ///     2014/05/07 集計項目に「休日出勤日数」を追加
        ///------------------------------------------------------------------------
        private string getTextPartData()
        {
            //出力文字列作成
            StringBuilder sb = new StringBuilder();
            sb.Append(pShainNo).Append(",");
            sb.Append(pShukkinDays.ToString()).Append(",");
            sb.Append(pKekkinDays.ToString()).Append(",");
            sb.Append(pShiteiKyuDays.ToString()).Append(",");
            sb.Append(pYuKyuDays.ToString()).Append(",");
            sb.Append(pkoukaDays.ToString()).Append(",");
            sb.Append(pShoteiTimes.ToString()).Append(",");
            sb.Append(pZan12Times.ToString()).Append(",");
            sb.Append(pZan125Times.ToString()).Append(",");
            sb.Append(pTokubetsuTeate.ToString()).Append(",");
            sb.Append(pShinyaTimes.ToString()).Append(",");
            sb.Append(pKyuShutsuTimes.ToString()).Append(",");
            sb.Append(pHanKyuDays.ToString()).Append(",");
            sb.Append(pKoshoDays.ToString()).Append(",");
            sb.Append(pTokubetsuKbn.ToString()).Append(",");
            sb.Append(pShoteiHanTimes.ToString()).Append(",");
            sb.Append(pKyuShutsuDays.ToString());   // 2014/05/07 追加

            return sb.ToString();
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     給与大臣受入データ初期化・社員, 出向社員 </summary>
        ///--------------------------------------------------------------
        private void InitOutRecShainShukko()
        {
            sShainNo = string.Empty;            // 社員番号
            sKyuShutsuTimes = global.flgOff;    // 休日出勤時間・分
            sKyuShutsuDays_d = global.flgOff;   // 休日出勤・代休あり日数
            sKyuShutsuDays = global.flgOff;     // 休日出勤・代休なし日数
            dummy1 = string.Empty;              // 使用しない
            sStockKyuuka = global.flgOff;       // ストック休暇
            sTokushuTeate = global.flgOff;      // 特殊作業時間・分
            sTokukhuDays_Non = global.flgOff;   // 特別休暇・無給
            sKyushokuDays = global.flgOff;      // 休職日数
            sShukkinDays = global.flgOff;       // 出勤日数
            sKyujitsuDays = global.flgOff;      // 休日日数
            sHanKyuDays = global.flgOff;        // 半休日数
            sYuKyuDays = global.flgOff;         // 有休日数
            sDaiKyuDays = global.flgOff;        // 代休日数
            sTokuKyuDays = global.flgOff;       // 特別休暇日数
            sKoshoDays = global.flgOff;         // 公傷日数
            sShuchoDays = global.flgOff;        // 出張日数
            sKekkinDays = global.flgOff;        // 欠勤日数
            sZanTimes = global.flgOff;          // 早出残業時間・分
            dummy2 = string.Empty;              // 使用しない
            sDaiKyuTimes = global.flgOff;       // 代休時間・分
            sShinyaTimes = global.flgOff;       // 深夜時間・分
            sYobi1 = global.flgOff;             // 呼出・昼回数
            sYobi2 = global.flgOff;             // 呼出・夜回数
            sAsaCount = global.flgOff;          // 朝番回数
            sNakaCount = global.flgOff;         // 中番回数
            sYoruCount = global.flgOff;         // 夜番回数
            sChisouTimes = global.flgOff;       // 遅早時間・分
            sCutTimes = global.flgOff;          // カット時間・分
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     給与大臣受入データ初期化・パート </summary>
        ///     
        ///     2014/05/07 集計項目「休日出勤日数」を追加
        ///--------------------------------------------------------------
        private void InitOutRecPart()
        {
            pShainNo = string.Empty;            // 汎用データ：社員番号
            pShukkinDays = global.flgOff;       // 出勤日数
            pKekkinDays = global.flgOff;        // 欠勤日数
            pShiteiKyuDays = global.flgOff;     // 指定休日数
            pYuKyuDays = global.flgOff;         // 有休日数
            pkoukaDays = global.flgOff;         // 公暇日数
            pShoteiTimes = global.flgOff;       // 労働時間・分：「出勤」の日の「所定内・休日」合計
            pZan12Times = global.flgOff;        // 残業1.2・分
            pZan125Times = global.flgOff;       // 残業1.25・分
            pTokubetsuTeate = global.flgOff;    // 特別手当・分
            pShinyaTimes = global.flgOff;       // 深夜時間・分
            pKyuShutsuTimes = global.flgOff;    // 休日出勤時間・分：「休日出勤・代休無し」日の「所定内・休日」合計
            pHanKyuDays = global.flgOff;        // 半休日数
            pKoshoDays = global.flgOff;         // 公傷日数
            pTokubetsuKbn = global.flgOff;      // 特別手当区分
            pShoteiHanTimes = global.flgOff;    // 労働時間・半休：「有休・半日」日の「所定内・休日」合計・分
            pKyuShutsuDays = global.flgOff;     // 休日出勤日数 2014/05/07 追加
        }
    }
}
