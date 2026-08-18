using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SZOK_OCR.Common
{
    public class ClsMaster : ClsSqlServerConnect, IMaster
    {
        public ClsMaster(string server, string login, string pw, string database) : base(server, login, pw, database)
        {
        }

        public SqlConnection OpenConnection()
        {
            SqlConnection conn = new SqlConnection(sqlBuilder.ConnectionString);

            try
            {
                if (conn.State == System.Data.ConnectionState.Closed)
                {
                    conn.Open();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return conn;
        }

        // SQL接続を閉じる
        public void CloseConnection(SqlConnection conn)
        {
            try
            {
                if (conn.State == System.Data.ConnectionState.Open)
                {
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public T GetData<T>(string sCode, SqlConnection conn)
        {
            // 環境設定のとき
            if (typeof(T) == typeof(TblConfig))
            {
                return (T)(Object)GetClsConfigData(Utility.StrtoInt(sCode), conn);
            }

            //// 商品マスターのとき
            //if (typeof(T) == typeof(ClsShohinData))
            //{
            //    return (T)(Object)GetClsShohinData(sCode);
            //}

            //// 得意先マスターのとき
            //if (typeof(T) == typeof(ClsTokuisakiData))
            //{
            //    return (T)(Object)GetClsTokuisakiData(sCode);
            //}

            //// 発注書ヘッダのとき
            //if (typeof(T) == typeof(ClsOrderData.Head))
            //{
            //    return (T)(Object)GetOrderHeadData(sCode);
            //}

            //// 仮伝票番号テーブルのとき
            //if (typeof(T) == typeof(ClsKariDenNum))
            //{
            //    return (T)(Object)GetClsKariDenNum(sCode);
            //}

            MessageBox.Show("Invalid Data Class");
            return default(T);
        }

        public void Insert<T>(T cls, SqlConnection conn)
        {
            // SCANDATAのとき
            if (typeof(T) == typeof(TblScandata))
            {
                Insert((TblScandata)(object)(cls), conn);
            }
        }

        public void UpDate<T>(T cls, SqlConnection conn)
        {
            // 環境設定のとき
            if (typeof(T) == typeof(TblConfig))
            {
                UpDate((TblConfig)(object)(cls), conn);
            }
        }

        /// <summary>
        /// 環境設定テーブルのデータを取得する：2026/08/18
        /// </summary>
        /// <param name="id">ID</param>
        /// <param name="conn">SQL接続オブジェクト</param>
        /// <returns>TblConfigクラス</returns>
        private TblConfig GetClsConfigData(int id, SqlConnection conn)
        {
            TblConfig configData = null;

            try
            {
                string sql = "SELECT * FROM 環境設定 WHERE ID = @ID";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@ID", id);

                    using (var dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            configData = new TblConfig
                            {
                                ID = id,
                                Year = 0,
                                Month = 0,
                                DataSaveMonth = Utility.StrtoInt(Utility.NulltoStr(dr["データ保存月数"])),
                                ZipCodePath = Utility.NulltoStr(dr["郵便番号データパス"])
                            };
                        }
                    }
                }

                return configData;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return configData;
            }
        }

        /// <summary>
        /// 環境設定テーブルのデータを更新する：2026/08/18
        /// </summary>
        /// <param name="configData">TblConfigクラス</param>
        /// <param name="conn">SQL接続オブジェクト</param>
        private void UpDate(TblConfig configData, SqlConnection conn)
        {
            try
            {
                string sql = "UPDATE 環境設定 SET " +
                             "年 = @year, 月 = @month, データ保存月数 = @dataSave, 郵便番号データパス = @zipPath, " +
                             "更新年月日 = @upDate " +
                             "WHERE ID = @ID";

                using (SqlCommand com = new SqlCommand(sql, conn))
                {
                    com.Parameters.AddWithValue("@year", configData.Year);
                    com.Parameters.AddWithValue("@month", configData.Month);
                    com.Parameters.AddWithValue("@dataSave", configData.DataSaveMonth);
                    com.Parameters.AddWithValue("@zipPath", configData.ZipCodePath);
                    com.Parameters.AddWithValue("@upDate", DateTime.Now.ToString());
                    com.Parameters.AddWithValue("@ID", configData.ID);
                    com.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// SCANDATAテーブルにデータを挿入する：2026/08/18
        /// </summary>
        /// <param name="scandata">SCAN_DATAクラス</param>
        /// <param name="conn">SQL接続オブジェクト</param>
        public void Insert(TblScandata scandata, SqlConnection conn)
        {
            string sql = "INSERT INTO SCAN_DATA (データ区分, 画像名, 登録年, 登録月, 登録日, 登録番号," +
                "車体番号, メーカー, 塗色, 車種, 郵便番号1, 郵便番号2, 車両番号1, 車両番号2, 車名, 住所漢字," +
                "住所1, 住所2, 氏名, TEL携帯, TEL携帯2, TEL携帯3, PC名, CSV作成日, 備考, 更新年月日, ラベル, 処理担当者) " +
                "values (@datakbn, @imageName, @Year, @Month, @Day, @number, @VehicleIdentificationNumber," +
                "@Maker, @Color, @CarModel, @ZipCode1, @ZipCode2, @VehicleNumber1, @VehicleNumber2, @CarName, @AddressKanji," +
                "@Address1, @Address2, @Name, @Mobile1, @Mobile2, @Mobile3, @PC, @CsvCreationDate, @Memo, @UpDate, @Label, @Person)";

            using (SqlCommand com = new SqlCommand(sql, conn))
            {
                com.Parameters.AddWithValue("@datakbn", scandata.DataCategory);
                com.Parameters.AddWithValue("@imageName", scandata.ImageFileName);
                com.Parameters.AddWithValue("@Year", scandata.AddYear);
                com.Parameters.AddWithValue("@Month", scandata.AddMonth);
                com.Parameters.AddWithValue("@Day", scandata.AddDay);
                com.Parameters.AddWithValue("@number", scandata.Number);
                com.Parameters.AddWithValue("@VehicleIdentificationNumber", scandata.VehicleIdentificationNumber);
                com.Parameters.AddWithValue("@Maker", scandata.Maker);
                com.Parameters.AddWithValue("@Color", scandata.Color);
                com.Parameters.AddWithValue("@CarModel", scandata.CarModel);
                com.Parameters.AddWithValue("@ZipCode1", scandata.ZipCode1);
                com.Parameters.AddWithValue("@ZipCode2", scandata.ZipCode2);
                com.Parameters.AddWithValue("@VehicleNumber1", scandata.VehicleNumber1);
                com.Parameters.AddWithValue("@VehicleNumber2", scandata.VehicleNumber2);
                com.Parameters.AddWithValue("@CarName", scandata.CarName);
                com.Parameters.AddWithValue("@AddressKanji", scandata.AddressKanji);
                com.Parameters.AddWithValue("@Address1", scandata.Address1);
                com.Parameters.AddWithValue("@Address2", scandata.Address2);
                com.Parameters.AddWithValue("@Name", scandata.Name);
                com.Parameters.AddWithValue("@Mobile1", scandata.Mobile1);
                com.Parameters.AddWithValue("@Mobile2", scandata.Mobile2);
                com.Parameters.AddWithValue("@Mobile3", scandata.Mobile3);
                com.Parameters.AddWithValue("@PC", scandata.PC);
                com.Parameters.AddWithValue("@CsvCreationDate", scandata.CsvCreationDate);
                com.Parameters.AddWithValue("@Memo", scandata.Memo);
                com.Parameters.AddWithValue("@UpDate", DateTime.Now.ToString());
                com.Parameters.AddWithValue("@Label", scandata.Label);
                com.Parameters.AddWithValue("@Person", scandata.Person);
                com.ExecuteNonQuery();
            }
        }
    }
}
