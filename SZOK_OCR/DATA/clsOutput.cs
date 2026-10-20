using Microsoft.Win32;
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
            // コメント化： 2026/09/03
            //// データ読み込み
            ////adp.Fill(dts.防犯登録データ);    // コメント化 2019/06/25
            //adp.FillByCSV(dts.防犯登録データ); // 2019/06/25 絞込

            //cAdp.Fill(dts.CSV作成履歴);

            // 郵便番号ＣＳＶデータを配列に読み込む
            Utility.zipCsvLoad(ref zipArray);
        }

        // コメント化： 2026/09/03
        //cardDataSet dts = new cardDataSet();
        //cardDataSetTableAdapters.防犯登録データTableAdapter adp = new cardDataSetTableAdapters.防犯登録データTableAdapter();
        //cardDataSetTableAdapters.CSV作成履歴TableAdapter cAdp = new cardDataSetTableAdapters.CSV作成履歴TableAdapter();

        string[] zipArray = null;   // 郵便番号配列

        /// <summary>自転車防犯登録カードデータ CSVファイル出力：2026/09/03　</summary>
        /// <param name="csv_Category">CSVカテゴリー</param>
        /// <returns>出力件数</returns>
        public int SaveCycleCsv(string csv_Category)
        {
            // 出力配列
            string[] arrayCsv = null;

            StringBuilder sb = new StringBuilder();
            int cnt = 0;

            string add1 = "";
            string add2 = "";

            var master = new ClsMaster(Properties.Settings.Default.sServerName, Properties.Settings.Default.sLogin, Properties.Settings.Default.sPass, Properties.Settings.Default.sDatabase);
            
            List<TblRegistrationCard> regis = new List<TblRegistrationCard>();

            if (csv_Category == global.CSV_TblRegistrationCard)
            {
                // 防犯登録データ
                regis = master.ReadCsvCreationData(global.DATA_CYCLE);
            }
            else if (csv_Category == global.CSV_ChangeData)
            {
                // 変更データ
                regis = master.ChangeCsvCreationData(global.DATA_CYCLE);
            }
            else if (csv_Category == global.CSV_Erasure)
            {
                // 抹消データ
                regis = master.ErasureCsvCreationData(global.DATA_CYCLE);
            }

            //var regis = master.ReadCsvCreationData(global.DATA_CYCLE);

            foreach (var t in regis)
            {
                cnt++;

                sb.Clear();
                sb.Append(global.DATA_CPA).Append(",");
                sb.Append(t.Number.Replace(global.DATA_CPA, "").Trim()).Append(",");
                sb.Append(t.VehicleIdentificationNumber).Append(",");

                // 西暦4ケタに変更 : 2016/03/08
                sb.Append("20" + t.AddYear + t.AddMonth.PadLeft(2, '0') + t.AddDay.PadLeft(2, '0')).Append(",");
                sb.Append(t.Maker.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.Color.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.CarModel.ToString().PadLeft(2, '0').Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");

                // 郵便番号住所とそれ以下の住所を各々取得する : 2016/03/08
                add1 = "";
                add2 = "";

                // 2016/05/26
                if (t.Address1.Contains(global.KENGAI_ADD))
                {
                    // 「ｹﾝｶﾞｲ」のとき
                    add1 = global.KENGAI_ADD;
                    add2 = t.Address1.Replace(global.KENGAI_ADD, "").Trim();
                }
                else
                {
                    // getAddressSplit(out add1, out add2, t.住所1); // 郵便番号住所とそれ以下の住所に分割する
                    getAddressSplitCity(out add1, out add2, t.Address1); // 市区町村とそれ以下の住所に分割する 2016/06/08
                }

                sb.Append(add1.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");    // 市区町村とそれ以下の住所 : 2016/03/08
                sb.Append(add2.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");    // 以下の住所 : 2016/03/08

                sb.Append(t.Name.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.Mobile1.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.Mobile2.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.Mobile3.Replace("\r", "").Replace("\n", "").Replace(",", ""));

                // 配列にセット
                Array.Resize(ref arrayCsv, cnt);        // 配列のサイズ拡張
                arrayCsv[cnt - 1] = sb.ToString();      // 文字列のセット
            }

            if (cnt > 0)
            {
                // CSVファイル名を設定：2026/10/20
                string csvFileName = "";

                if (csv_Category == global.CSV_TblRegistrationCard)
                {
                    csvFileName = global.CSV_CYCLE;
                }
                else if (csv_Category == global.CSV_ChangeData || csv_Category == global.CSV_Erasure)
                {
                    csvFileName = global.CSV_CYCLE + "_" + csv_Category;
                }
                else
                {
                    csvFileName = global.CSV_CYCLE;
                }

                // CSVファイル出力
                Utility.CsvFileWrite(global.cnfPath, arrayCsv, csvFileName);

                // CSV作成日を登録：2026/09/03
                if (csv_Category == global.CSV_TblRegistrationCard)
                {
                    // 防犯登録データにCSV作成日を書き込み：2026/09/03
                    foreach (var t in regis)
                    {
                        master.UpDateRegisCsvCreation(t.ID);
                    }
                }

                // 変更届データのCSV作成日：2026/10/20
                if (csv_Category == global.CSV_ChangeData)
                {
                    // 変更届データに変更届CSV作成日を書き込み：2026/10/20
                    foreach (var t in regis)
                    {
                        master.UpDateChangeCsvCreation(t.Number);
                    }
                }

                // 抹消データのCSV作成日：2026/10/20
                if (csv_Category == global.CSV_Erasure)
                {
                    // 抹消データに抹消CSV作成日を書き込み：2026/10/20
                    foreach (var t in regis)
                    {
                        master.UpDateErasureCsvCreation(t.ID);
                    }
                }

                // CSV作成履歴クラス作成：2026/09/03
                var csvCreationHistory = new TblCSVCreationHistory
                {
                    CreationDate = DateTime.Now,
                    Outputs = cnt,
                    PC = Environment.MachineName,
                    //Memo = global.CSV_CYCLE   // 2026/10/20 コメント化
                    Memo = csvFileName  // 2026/10/20
                };

                // CSV作成履歴を登録：2026/09/03
                master.Insert(csvCreationHistory);
            }

            return cnt;
        }

        /// <summary>原付防犯登録カードデータ CSVファイル出力： 2026/09/04　</summary>
        /// <param name="csv_Category">CSVカテゴリー</param>
        /// <returns>出力件数</returns>
        public int SaveAutoCsv(string csv_Category)
        {
            // 出力配列
            string[] arrayCsv = null;

            StringBuilder sb = new StringBuilder();
            int cnt = 0;

            string add1 = "";
            string add2 = "";

            var master = new ClsMaster(Properties.Settings.Default.sServerName, Properties.Settings.Default.sLogin, Properties.Settings.Default.sPass, Properties.Settings.Default.sDatabase);
            //var regis = master.ReadCsvCreationData(global.DATA_AUTO);

            List<TblRegistrationCard> regis = new List<TblRegistrationCard>();

            if (csv_Category == global.CSV_TblRegistrationCard)
            {
                // 防犯登録データ
                regis = master.ReadCsvCreationData(global.DATA_AUTO);
            }
            else if (csv_Category == global.CSV_ChangeData)
            {
                // 変更データ
                regis = master.ChangeCsvCreationData(global.DATA_AUTO);
            }
            else if (csv_Category == global.CSV_Erasure)
            {
                // 抹消データ
                regis = master.ErasureCsvCreationData(global.DATA_AUTO);
            }


            foreach (var t in regis)
            {
                cnt++;
                
                sb.Clear();
                sb.Append(global.DATA_CPA).Append(",");
                sb.Append(t.Number.Replace(global.DATA_CPA, "").Trim()).Append(",");
                sb.Append(t.VehicleIdentificationNumber).Append(",");
                sb.Append(t.VehicleNumber1).Append(",");
                sb.Append(t.VehicleNumber2).Append(",");

                // 西暦4ケタに変更 : 2016/03/08
                sb.Append("20" + t.AddYear + t.AddMonth.PadLeft(2, '0') + t.AddDay.PadLeft(2, '0')).Append(",");
                sb.Append(t.Maker.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.CarName.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.Color.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.CarModel.ToString().PadLeft(2, '0').Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");

                // 郵便番号住所とそれ以下の住所を各々取得する : 2016/03/08
                add1 = "";
                add2 = "";

                // 2016/05/26
                if (t.Address1.Contains(global.KENGAI_ADD))
                {
                    // 「ｹﾝｶﾞｲ」のとき
                    add1 = global.KENGAI_ADD;
                    add2 = t.Address1.Replace(global.KENGAI_ADD, "").Trim();
                }
                else
                {
                    //getAddressSplit(out add1, out add2, t.住所1);     // 郵便番号住所とそれ以下の住所に分割する
                    getAddressSplitCity(out add1, out add2, t.Address1); // 市区町村とそれ以下の住所に分割する 2016/06/08
                }

                sb.Append(add1.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");    // 郵便番号住所 : 2016/03/08
                sb.Append(add2.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");    // 以下の住所 : 2016/03/08

                sb.Append(t.Name.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.Mobile1.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.Mobile2.Replace("\r", "").Replace("\n", "").Replace(",", "")).Append(",");
                sb.Append(t.Mobile3.Replace("\r", "").Replace("\n", "").Replace(",", ""));

                // 配列にセット
                Array.Resize(ref arrayCsv, cnt);        // 配列のサイズ拡張
                arrayCsv[cnt - 1] = sb.ToString();      // 文字列のセット
            }

            if (cnt > 0)
            {
                // CSVファイル名を設定：2026/10/20
                string csvFileName = "";

                if (csv_Category == global.CSV_TblRegistrationCard)
                {
                    csvFileName = global.CSV_AUTO;
                }
                else if (csv_Category == global.CSV_ChangeData || csv_Category == global.CSV_Erasure)
                {
                    csvFileName = global.CSV_AUTO + "_" + csv_Category;
                }
                else
                {
                    csvFileName = global.CSV_AUTO;
                }

                // CSVファイル出力
                Utility.CsvFileWrite(global.cnfPath, arrayCsv, csvFileName);

                // 防犯登録データにCSV作成日を登録：2026/09/03
                //foreach (var t in regis)
                //{
                //    master.UpDateRegisCsvCreation(t.ID);
                //}

                // CSV作成日を登録：2026/09/03
                if (csv_Category == global.CSV_TblRegistrationCard)
                {
                    // 防犯登録データにCSV作成日を書き込み：2026/09/03
                    foreach (var t in regis)
                    {
                        master.UpDateRegisCsvCreation(t.ID);
                    }
                }

                // 変更届データのCSV作成日：2026/10/20
                if (csv_Category == global.CSV_ChangeData)
                {
                    // 変更届データに変更届CSV作成日を書き込み：2026/10/20
                    foreach (var t in regis)
                    {
                        master.UpDateChangeCsvCreation(t.Number);
                    }
                }

                // 抹消データのCSV作成日：2026/10/20
                if (csv_Category == global.CSV_Erasure)
                {
                    // 抹消データに抹消CSV作成日を書き込み：2026/10/20
                    foreach (var t in regis)
                    {
                        master.UpDateErasureCsvCreation(t.ID);
                    }
                }

                // CSV作成履歴クラス作成：2026/09/04
                var csvCreationHistory = new TblCSVCreationHistory
                {
                    CreationDate = DateTime.Now,
                    Outputs = cnt,
                    PC = Environment.MachineName,
                    //Memo = global.CSV_AUTO   // 2026/10/20 コメント化
                    Memo = csvFileName  // 2026/10/20
                };

                // CSV作成履歴を登録：2026/09/03
                master.Insert(csvCreationHistory);
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
            //var s = dts.CSV作成履歴.NewCSV作成履歴Row();
            //s.作成年月日 = DateTime.Now;
            //s.出力件数 = cnt;
            //s.PC名 = Environment.MachineName;
            //s.摘要 = sTekiyo;
            //dts.CSV作成履歴.AddCSV作成履歴Row(s);

            //// データベース更新
            //cAdp.Update(dts.CSV作成履歴);
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
