using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;

namespace SZOK_OCR.Common
{
    public class ClsMaster : ClsSqlServerConnect, IMaster
    {
        public ClsMaster(string server, string login, string pw, string database) : base(server, login, pw, database)
        {
        }

        /// <summary>
        /// SQL接続を開く
        /// </summary>
        /// <returns>SqlConnectionオブジェクト</returns>
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

        /// <summary>
        /// SQL接続を閉じる
        /// </summary>
        /// <param name="conn">SqlConnectionオブジェクト</param>
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

        /// <summary>
        /// 指定されたコードに基づいてデータを取得する
        /// </summary>
        /// <typeparam name="T">データの型</typeparam>
        /// <param name="sCode">コード</param>
        /// <param name="conn">SQL接続オブジェクト</param>
        /// <returns>指定された型のデータ</returns>
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

        /// <summary>
        /// 指定されたデータを挿入する
        /// </summary>
        /// <typeparam name="T">データの型</typeparam>
        /// <param name="cls">挿入するデータ</param>
        /// <param name="conn">SQL接続オブジェクト</param>
        public void Insert<T>(T cls, SqlConnection conn)
        {
            // SCANDATAのとき
            if (typeof(T) == typeof(TblScandata))
            {
                Insert((TblScandata)(object)(cls), conn);
            }
        }

        /// <summary>
        /// 指定されたデータを更新する
        /// </summary>
        /// <typeparam name="T">データの型</typeparam>
        /// <param name="cls">更新するデータ</param>
        /// <param name="conn">SQL接続オブジェクト</param>
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

        public List<T> Read<T>(ScandataParameter param, SqlConnection conn)
        {
            // SCANDATAのとき
            if (typeof(T) == typeof(TblScandata))
            {
                return Read(param, conn) as List<T>;
            }

            //List<T> resultList = new List<T>();
            //using (SqlCommand cmd = new SqlCommand(sql, conn))
            //{
            //    using (SqlDataReader reader = cmd.ExecuteReader())
            //    {
            //        while (reader.Read())
            //        {
            //            T obj = Activator.CreateInstance<T>();
            //            for (int i = 0; i < reader.FieldCount; i++)
            //            {
            //                var property = typeof(T).GetProperty(reader.GetName(i));
            //                if (property != null && !reader.IsDBNull(i))
            //                {
            //                    property.SetValue(obj, reader.GetValue(i));
            //                }
            //            }
            //            resultList.Add(obj);
            //        }
            //    }
            //}
            //return resultList;

            MessageBox.Show("Invalid Data Class");
            return null;
        }


        public List<TblScandata> Read(ScandataParameter param, SqlConnection conn)
        {
            var lines = new List<TblScandata>();

            try
            {
                string sql = "SELECT * FROM SCAN_DATA " +
                    "WHERE (" +
                    "(@DataCategory IS NULL OR データ区分 = @DataCategory) AND " +
                    "(NULLIF(@AddYear, '') IS NULL OR 登録年 = @AddYear) AND " +
                    "(NULLIF(@AddMonth, '') IS NULL OR 登録月 = @AddMonth) AND " +
                    "(NULLIF(@AddDay, '') IS NULL OR 登録日 = @AddDay) AND " +
                    "(NULLIF(@Number, '') IS NULL OR 登録番号 LIKE '%' + @Number + '%') AND " +
                    "(NULLIF(@VehicleIdentificationNumber, '') IS NULL OR 車体番号 LIKE '%' + @VehicleIdentificationNumber + '%') AND " +
                    "(NULLIF(@Maker, '') IS NULL OR メーカー LIKE '%' + @Maker + '%') AND " +
                    "(NULLIF(@Color, '') IS NULL OR 塗色 LIKE '%' + @Color + '%') AND " +
                    "(@CarModel IS NULL OR 車種 = @CarModel) AND " +
                    "(NULLIF(@ZipCode1, '') IS NULL OR 郵便番号1 LIKE '%' + @ZipCode1 + '%') AND " +
                    "(NULLIF(@ZipCode2, '') IS NULL OR 郵便番号2 LIKE '%' + @ZipCode2 + '%') AND " +
                    "(NULLIF(@Address1, '') IS NULL OR 住所1 LIKE '%' + @Address1 + '%') AND " +
                    "(NULLIF(@Name, '') IS NULL OR 氏名 LIKE '%' + @Name + '%') AND " +
                    "(NULLIF(@Mobile1, '') IS NULL OR TEL携帯 LIKE '%' + @Mobile1 + '%') AND " +
                    "(NULLIF(@Mobile2, '') IS NULL OR TEL携帯2 LIKE '%' + @Mobile2 + '%') AND " +
                    "(NULLIF(@Mobile3, '') IS NULL OR TEL携帯3 LIKE '%' + @Mobile3 + '%') AND " +
                    "(NULLIF(@Label, '') IS NULL OR ラベル LIKE '%' + @Label + '%') AND " +
                    "(NULLIF(@Person, '') IS NULL OR 処理担当者 LIKE '%' + @Person + '%')) " +
                    "ORDER BY 登録番号";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    var dataCategoryParameter = cmd.Parameters.Add("@DataCategory", System.Data.SqlDbType.Int);
                    dataCategoryParameter.Value = param.DataCategory.HasValue ? (object)param.DataCategory.Value : DBNull.Value;
                    //cmd.Parameters.AddWithValue("@DataCategory", param.DataCategory);
                    cmd.Parameters.AddWithValue("@AddYear", param.AddYear);
                    cmd.Parameters.AddWithValue("@AddMonth", param.AddMonth);
                    cmd.Parameters.AddWithValue("@AddDay", param.AddDay);
                    cmd.Parameters.AddWithValue("@Number", param.Number);
                    cmd.Parameters.AddWithValue("@VehicleIdentificationNumber", param.VehicleIdentificationNumber);
                    cmd.Parameters.AddWithValue("@Maker", param.Maker);
                    cmd.Parameters.AddWithValue("@Color", param.Color);


                    var CarModelParameter = cmd.Parameters.Add("@CarModel", System.Data.SqlDbType.Int);
                    CarModelParameter.Value = param.CarModel.HasValue ? (object)param.CarModel.Value : DBNull.Value;
                    //cmd.Parameters.AddWithValue("@CarModel", Utility.StrtoInt(param.CarModel));
                    cmd.Parameters.AddWithValue("@ZipCode1", param.ZipCode1);
                    cmd.Parameters.AddWithValue("@ZipCode2", param.ZipCode2);
                    cmd.Parameters.AddWithValue("@Address1", param.Address1);
                    cmd.Parameters.AddWithValue("@Name", param.Name);
                    cmd.Parameters.AddWithValue("@Mobile1", param.Mobile1);
                    cmd.Parameters.AddWithValue("@Mobile2", param.Mobile2);
                    cmd.Parameters.AddWithValue("@Mobile3", param.Mobile3);
                    cmd.Parameters.AddWithValue("@Label", param.Label);
                    cmd.Parameters.AddWithValue("@Person", param.Person);

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            TblScandata scandata = new TblScandata
                            {
                                ID = Utility.StrtoInt(dr["ID"].ToString()),
                                DataCategory = Utility.StrtoInt(Utility.NulltoStr(dr["データ区分"])),
                                ImageFileName = Utility.NulltoStr(dr["画像名"]),
                                AddYear = Utility.NulltoStr(dr["登録年"]),
                                AddMonth = Utility.NulltoStr(dr["登録月"]),
                                AddDay = Utility.NulltoStr(dr["登録日"]),
                                Number = Utility.NulltoStr(dr["登録番号"]),
                                VehicleIdentificationNumber = Utility.NulltoStr(dr["車体番号"]),
                                Maker = Utility.NulltoStr(dr["メーカー"]),
                                Color = Utility.NulltoStr(dr["塗色"]),
                                CarModel = Utility.StrtoInt(Utility.NulltoStr(dr["車種"])),
                                ZipCode1 = Utility.NulltoStr(dr["郵便番号1"]),
                                ZipCode2 = Utility.NulltoStr(dr["郵便番号2"]),
                                VehicleNumber1 = Utility.NulltoStr(dr["車両番号1"]),
                                VehicleNumber2 = Utility.NulltoStr(dr["車両番号2"]),
                                CarName = Utility.NulltoStr(dr["車名"]),
                                Address1 = Utility.NulltoStr(dr["住所1"]),
                                Address2 = Utility.NulltoStr(dr["住所2"]),
                                Name = Utility.NulltoStr(dr["氏名"]),
                                Mobile1 = Utility.NulltoStr(dr["TEL携帯"]),
                                Mobile2 = Utility.NulltoStr(dr["TEL携帯2"]),
                                Mobile3 = Utility.NulltoStr(dr["TEL携帯3"]),
                                PC = Utility.NulltoStr(dr["PC名"]),
                                CsvCreationDate = Utility.NulltoStr(dr["CSV作成日"]),
                                Memo = Utility.NulltoStr(dr["備考"]),
                                UpDate = DateTime.Parse(Utility.NulltoStr(dr["更新年月日"])),
                                Label = Utility.NulltoStr(dr["ラベル"]),
                                Person = Utility.NulltoStr(dr["処理担当者"])
                            };

                            lines.Add(scandata);
                        }
                    }
                }
                return lines;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return lines;
            }
        }
    }
}
