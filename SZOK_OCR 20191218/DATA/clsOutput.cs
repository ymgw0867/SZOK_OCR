using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SZOK_OCR.Common;

namespace SZOK_OCR.DATA
{
    class clsOutput
    {
        public clsOutput()
        {
            // データ読み込み
            //adp.Fill(dts.防犯登録データ);    // コメント化 2019/06/25
            adp.FillByCSV(dts.防犯登録データ); // 2019/06/25 絞込

            cAdp.Fill(dts.CSV作成履歴);

            // 郵便番号ＣＳＶデータを配列に読み込む
            Utility.zipCsvLoad(ref zipArray);
        }

        cardDataSet dts = new cardDataSet();
        cardDataSetTableAdapters.防犯登録データTableAdapter adp = new cardDataSetTableAdapters.防犯登録データTableAdapter();
        cardDataSetTableAdapters.CSV作成履歴TableAdapter cAdp = new cardDataSetTableAdapters.CSV作成履歴TableAdapter();
        
        string[] zipArray = null;   // 郵便番号配列

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     自転車防犯登録カードデータ CSVファイル出力　</summary>
        /// <returns>
        ///     出力件数</returns>
        /// -------------------------------------------------------------------------
        public int saveCycleCsv()
        {
            // 出力配列
            string[] arrayCsv = null;
            
            StringBuilder sb = new StringBuilder();
            int cnt = 0;

            string add1 = "";
            string add2 = "";

            foreach (var t in dts.防犯登録データ.Where(a => a.データ区分 == global.DATA_CYCLE && 
                                                     (a.IsCSV作成日Null() || a.CSV作成日 == string.Empty) && 
                                                     (a.Is除外Null() || a.除外 == global.flgOff))
                                               .OrderBy(a => a.登録番号))
            {
                cnt++;

                sb.Clear();
                sb.Append(global.DATA_CPA).Append(",");
                sb.Append(t.登録番号.Replace(global.DATA_CPA, "").Trim()).Append(",");
                sb.Append(t.車体番号).Append(",");

                // 西暦4ケタに変更 : 2016/03/08
                sb.Append("20" + t.登録年 + t.登録月.PadLeft(2, '0') + t.登録日.PadLeft(2, '0')).Append(",");
                sb.Append(t.メーカー.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.塗色.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.車種.ToString().PadLeft(2, '0').Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");

                // 郵便番号住所とそれ以下の住所を各々取得する : 2016/03/08
                add1 = "";
                add2 = "";

                // 2016/05/26
                if (t.住所1.Contains(global.KENGAI_ADD))
                {
                    // 「ｹﾝｶﾞｲ」のとき
                    add1 = global.KENGAI_ADD;
                    add2 = t.住所1.Replace(global.KENGAI_ADD, "").Trim();
                }
                else
                {
                    // getAddressSplit(out add1, out add2, t.住所1); // 郵便番号住所とそれ以下の住所に分割する
                    getAddressSplitCity(out add1, out add2, t.住所1); // 市区町村とそれ以下の住所に分割する 2016/06/08
                }

                sb.Append(add1.Replace("\r", "").Replace("\n", "").Replace(",","")).Append(",");    // 市区町村とそれ以下の住所 : 2016/03/08
                sb.Append(add2.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");    // 以下の住所 : 2016/03/08

                sb.Append(t.氏名.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.TEL携帯.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.TEL携帯2.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.TEL携帯3.Replace("\r", "").Replace("\n", "").Replace(",", ""));

                // 配列にセット
                Array.Resize(ref arrayCsv, cnt);        // 配列のサイズ拡張
                arrayCsv[cnt - 1] = sb.ToString();      // 文字列のセット

                // CSV作成日を登録
                t.CSV作成日 = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0') + 
                             " " + DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Minute.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Second.ToString().PadLeft(2, '0');
                t.更新年月日 = DateTime.Now;
            }

            if (cnt > 0)
            {
                // CSVファイル出力
                Utility.csvFileWrite(global.cnfPath, arrayCsv, global.CSV_CYCLE);

                // カードデータ更新
                adp.Update(dts.防犯登録データ);

                // 作成履歴
                csvRirekiUpdate(cnt, global.CSV_CYCLE);
            }

            return cnt;
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     原付防犯登録カードデータ CSVファイル出力　</summary>
        /// <returns>
        ///     出力件数</returns>
        /// -------------------------------------------------------------------------
        public int saveAutoCsv()
        {
            // 出力配列
            string[] arrayCsv = null;

            StringBuilder sb = new StringBuilder();
            int cnt = 0;

            string add1 = "";
            string add2 = "";

            foreach (var t in dts.防犯登録データ.Where(a => a.データ区分 == global.DATA_AUTO &&
                                                      (a.IsCSV作成日Null() || a.CSV作成日 == string.Empty) &&
                                                     (a.Is除外Null() || a.除外 == global.flgOff))
                                               .OrderBy(a => a.登録番号))
            {
                cnt++;

                sb.Clear();
                sb.Append(global.DATA_CPA).Append(",");
                sb.Append(t.登録番号.Replace(global.DATA_CPA, "").Trim()).Append(",");
                sb.Append(t.車体番号).Append(",");
                sb.Append(t.車両番号1).Append(",");
                sb.Append(t.車両番号2).Append(",");

                // 西暦4ケタに変更 : 2016/03/08
                sb.Append("20" + t.登録年 + t.登録月.PadLeft(2, '0') + t.登録日.PadLeft(2, '0')).Append(",");
                sb.Append(t.メーカー.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.車名.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.塗色.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.車種.ToString().PadLeft(2, '0').Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");

                // 郵便番号住所とそれ以下の住所を各々取得する : 2016/03/08
                add1 = "";
                add2 = "";

                // 2016/05/26
                if (t.住所1.Contains(global.KENGAI_ADD))
                {
                    // 「ｹﾝｶﾞｲ」のとき
                    add1 = global.KENGAI_ADD;
                    add2 = t.住所1.Replace(global.KENGAI_ADD, "").Trim();
                }
                else
                {
                    //getAddressSplit(out add1, out add2, t.住所1);     // 郵便番号住所とそれ以下の住所に分割する
                    getAddressSplitCity(out add1, out add2, t.住所1); // 市区町村とそれ以下の住所に分割する 2016/06/08
                }

                sb.Append(add1.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");    // 郵便番号住所 : 2016/03/08
                sb.Append(add2.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");    // 以下の住所 : 2016/03/08

                sb.Append(t.氏名.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.TEL携帯.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.TEL携帯2.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.TEL携帯3.Replace("\r", "").Replace("\n", "").Replace(",", ""));

                // 配列にセット
                Array.Resize(ref arrayCsv, cnt);        // 配列のサイズ拡張
                arrayCsv[cnt - 1] = sb.ToString();      // 文字列のセット

                // CSV作成日を登録
                t.CSV作成日 = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0') +
                             " " + DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Minute.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Second.ToString().PadLeft(2, '0');
                t.更新年月日 = DateTime.Now;
            }

            if (cnt > 0)
            {
                // CSVファイル出力
                Utility.csvFileWrite(global.cnfPath, arrayCsv, global.CSV_AUTO);

                // カードデータ更新
                adp.Update(dts.防犯登録データ);

                // 作成履歴
                csvRirekiUpdate(cnt, global.CSV_AUTO);
            }

            return cnt;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     静岡県警察本部用ＣＳＶデータ作成履歴を記録する </summary>
        /// <param name="cnt">
        ///     出力件数</param>
        /// <param name="sTekiyo">
        ///     自転車または原付</param>
        ///-----------------------------------------------------------------------
        private void csvRirekiUpdate(int cnt, string sTekiyo)
        {
            var s = dts.CSV作成履歴.NewCSV作成履歴Row();
            s.作成年月日 = DateTime.Now;
            s.出力件数 = cnt;
            s.PC名 = Environment.MachineName;
            s.摘要 = sTekiyo;
            dts.CSV作成履歴.AddCSV作成履歴Row(s);

            // データベース更新
            cAdp.Update(dts.CSV作成履歴);
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     住所を郵便番号住所とそれ以下の住所に分割する </summary>
        /// <param name="add1">
        ///     郵便番号住所</param>
        /// <param name="add2">
        ///     それ以下の住所</param>
        /// <param name="pAdd">
        ///     住所</param>
        ///---------------------------------------------------------------------
        private void getAddressSplit(out string add1, out string add2, string pAdd)
        {
            add1 = "";
            add2 = "";
            bool hit = false;

            foreach (var t in zipArray)
            {
                string[] zip = t.Split(',');

                string cAdd1 = Utility.strSmallTolarge((zip[4] + " " + zip[5]).Replace("\"", ""));
                string cAdd2 = Utility.strSmallTolarge((zip[4] + zip[5]).Replace("\"", ""));

                if (Utility.strSmallTolarge(pAdd).Contains(cAdd1))
                {
                    add1 = cAdd1;
                    add2 = pAdd.Replace(cAdd1, "").Trim();
                    hit = true;
                    break;
                }
                else if (Utility.strSmallTolarge(pAdd).Contains(cAdd2))
                {
                    add1 = cAdd2;
                    add2 = pAdd.Replace(cAdd2, "").Trim();
                    hit = true;
                    break;
                }
            }

            if (!hit)
            {
                add1 = "";
                add2 = pAdd;
            }
        }


        ///---------------------------------------------------------------------------
        /// <summary>
        ///     住所を市区町村とそれ以下の住所に分割する </summary>
        /// <param name="add1">
        ///     市区町村</param>
        /// <param name="add2">
        ///     それ以下の住所</param>
        /// <param name="pAdd">
        ///     住所</param>
        ///---------------------------------------------------------------------------
        private void getAddressSplitCity(out string add1, out string add2, string pAdd)
        {
            add1 = "";
            add2 = "";
            bool hit = false;

            foreach (var t in zipArray)
            {
                string[] zip = t.Split(',');

                string cAdd1 = Utility.strSmallTolarge(zip[4].Replace("\"", ""));
                //string cAdd2 = Utility.strSmallTolarge((zip[4] + zip[5]).Replace("\"", ""));

                if (Utility.strSmallTolarge(pAdd).Contains(cAdd1))
                {
                    add1 = cAdd1;
                    add2 = pAdd.Replace(cAdd1, "").Trim();
                    hit = true;
                    break;
                }
            }

            if (!hit)
            {
                add1 = "";
                add2 = pAdd;
            }
        }
    }
}
