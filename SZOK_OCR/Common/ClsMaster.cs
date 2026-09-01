using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SZOK_OCR.DATA;
using static ClosedXML.Excel.XLPredefinedFormat;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
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
        public T GetData<T>(string sCode)
        {
            // 環境設定のとき
            if (typeof(T) == typeof(TblConfig))
            {
                return (T)(Object)GetClsConfigData(Utility.StrtoInt(sCode));
            }

            // SCAN_DATAのとき
            if (typeof(T) == typeof(TblScandata))
            {
                return (T)(Object)GetClsScandata(Utility.StrtoInt(sCode));
            }

            // 防犯カードのとき
            if (typeof(T) == typeof(TblWorkcard))
            {
                return (T)(Object)GetClsWorkcard(Utility.StrtoInt(sCode));
            }

            // 防犯登録データのとき
            if (typeof(T) == typeof(TblRegistrationCard))
            {
                return (T)(Object)GetClsRegistrationCard(Utility.StrtoInt(sCode));
            }

            MessageBox.Show("Invalid Data Class");
            return default(T);
        }

        public void Insert<T>(List<T> cls)
        {
            // SCANDATAのとき
            if (typeof(T) == typeof(TblScandata))
            {
                Insert((List<TblScandata>)(object)(cls));
            }
        }

        /// <summary>
        /// 指定されたデータを更新する
        /// </summary>
        /// <typeparam name="T">データの型</typeparam>
        /// <param name="cls">更新するデータ</param>
        public void UpDate<T>(T cls)
        {
            // 環境設定のとき
            if (typeof(T) == typeof(TblConfig))
            {
                UpDate((TblConfig)(object)(cls));
            }

            // 防犯カードのとき
            if (typeof(T) == typeof(TblWorkcard))
            {
                UpDate((TblWorkcard)(object)(cls));
            }
            // 防犯登録データのとき
            if (typeof(T) == typeof(TblRegistrationCard))
            {
                UpDate((TblRegistrationCard)(object)(cls));
            }
        }

        /// <summary>
        /// 指定された型のデータ件数を取得する
        /// </summary>
        /// <typeparam name="T">データの型</typeparam>
        /// <returns>データ件数</returns>
        public int Count<T>()
        {
            // SCAN_DATAのとき
            if (typeof(T) == typeof(TblScandata))
            {
                string sql = "SELECT COUNT(*) FROM SCAN_DATA";
                return GetCount(sql);
            }

            MessageBox.Show("Invalid Data Class");
            return 0;
        }

        public int Count<T>(string id)
        {
            // 防犯カードのとき
            if (typeof(T) == typeof(TblWorkcard))
            {
                string sql = "SELECT COUNT(*) FROM 防犯カード WHERE PC名 = @val";
                return GetCount(sql, id);
            }

            MessageBox.Show("Invalid Data Class");
            return 0;
        }

        public int CountNumber<T>(string number)
        {
            // 防犯登録カードのとき
            if (typeof(T) == typeof(TblRegistrationCard))
            {
                string sql = "SELECT COUNT(*) FROM 防犯登録データ WHERE 登録番号 = @val";
                return GetCount(sql, number);
            }

            MessageBox.Show("Invalid Data Class");
            return 0;
        }

        public int CountNumber<T>(string number, string pcName)
        {
            // 防犯カードのとき
            if (typeof(T) == typeof(TblWorkcard))
            {
                string sql = "SELECT COUNT(*) FROM 防犯カード WHERE 登録番号 = @Number and PC名 = @val";
                return GetCount(sql, number, pcName);
            }

            MessageBox.Show("Invalid Data Class");
            return 0;
        }

        /// <summary>
        /// 環境設定テーブルのデータを取得する：2026/08/18
        /// </summary>
        /// <param name="id">ID</param>
        /// <param name="conn">SQL接続オブジェクト</param>
        /// <returns>TblConfigクラス</returns>
        private TblConfig GetClsConfigData(int id)
        {
            TblConfig configData = null;

            try
            {
                using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

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
        /// SCAN_DATAテーブルのデータを取得する：2026/08/18
        /// </summary>
        /// <param name="id">ID</param>
        /// <returns>TblScandataクラス</returns>
        private TblScandata GetClsScandata(int id)
        {
            TblScandata scandata = null;

            try
            {
                using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    string sql = "SELECT * FROM SCAN_DATA WHERE ID = @ID";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", id);

                        using (var dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                scandata = new TblScandata
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
                                    VehicleNumber1 = string.Empty,
                                    VehicleNumber2 = string.Empty,
                                    CarName = string.Empty,
                                    Address1 = Utility.NulltoStr(dr["住所1"]),
                                    Address2 = string.Empty,
                                    Name = Utility.NulltoStr(dr["氏名"]),
                                    Mobile1 = Utility.NulltoStr(dr["TEL携帯"]),
                                    Mobile2 = Utility.NulltoStr(dr["TEL携帯2"]),
                                    Mobile3 = Utility.NulltoStr(dr["TEL携帯3"]),
                                    PC = Utility.NulltoStr(dr["PC名"]),
                                    CsvCreationDate = Utility.NulltoStr(dr["CSV作成日"]),
                                    Memo = Utility.NulltoStr(dr["備考"]),
                                    UpDate = System.DateTime.Parse(Utility.NulltoStr(dr["更新年月日"])),
                                    Label = Utility.NulltoStr(dr["ラベル"]),
                                    Person = Utility.NulltoStr(dr["処理担当者"])
                                };
                            }
                        }
                    }
                }

                return scandata;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return scandata;
            }
        }

        /// <summary>
        /// 防犯カードテーブルのデータを取得する：2026/08/26
        /// </summary>
        /// <param name="id">防犯カードのID</param>
        /// <returns>TblWorkcardオブジェクト</returns>
        private TblWorkcard GetClsWorkcard(int id)
        {
            TblWorkcard workcard = null;

            try
            {
                using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    string sql = "SELECT * FROM 防犯カード WHERE ID = @ID";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", id);

                        using (var dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                workcard = new TblWorkcard
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
                                    VehicleNumber1 = string.Empty,
                                    VehicleNumber2 = string.Empty,
                                    CarName = string.Empty,
                                    AddressKanji = Utility.NulltoStr(dr["住所漢字"]),
                                    Address1 = Utility.NulltoStr(dr["住所1"]),
                                    Address2 = string.Empty,
                                    Name = Utility.NulltoStr(dr["氏名"]),
                                    Mobile1 = Utility.NulltoStr(dr["TEL携帯"]),
                                    Mobile2 = Utility.NulltoStr(dr["TEL携帯2"]),
                                    Mobile3 = Utility.NulltoStr(dr["TEL携帯3"]),
                                    PC = Utility.NulltoStr(dr["PC名"]),
                                    CsvCreationDate = Utility.NulltoStr(dr["CSV作成日"]),
                                    Memo = Utility.NulltoStr(dr["備考"]),
                                    UpDate = System.DateTime.Parse(Utility.NulltoStr(dr["更新年月日"])),
                                    Label = Utility.NulltoStr(dr["ラベル"]),
                                    Person = Utility.NulltoStr(dr["処理担当者"]),
                                    Confirmation = Utility.StrtoInt(Utility.NulltoStr(dr["確認"]))
                                };
                            }
                        }
                    }
                }

                return workcard;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return workcard;
            }
        }

        /// <summary>
        /// 防犯登録データテーブルのデータを取得する：2026/08/31
        /// </summary>
        /// <param name="id">防犯登録データのID</param>
        /// <returns>TblRegistrationCardオブジェクト</returns>
        private TblRegistrationCard GetClsRegistrationCard(int id)
        {
            TblRegistrationCard registrationCard = null;

            try
            {
                using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    string sql = "SELECT * FROM 防犯登録データ WHERE ID = @ID";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@ID", id);

                        using (var dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                registrationCard = new TblRegistrationCard
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
                                    VehicleNumber1 = string.Empty,
                                    VehicleNumber2 = string.Empty,
                                    CarName = string.Empty,
                                    AddressKanji = Utility.NulltoStr(dr["住所漢字"]),
                                    Address1 = Utility.NulltoStr(dr["住所1"]),
                                    Address2 = string.Empty,
                                    Name = Utility.NulltoStr(dr["氏名"]),
                                    Mobile1 = Utility.NulltoStr(dr["TEL携帯"]),
                                    Mobile2 = Utility.NulltoStr(dr["TEL携帯2"]),
                                    Mobile3 = Utility.NulltoStr(dr["TEL携帯3"]),
                                    PC = Utility.NulltoStr(dr["PC名"]),
                                    CsvCreationDate = Utility.NulltoStr(dr["CSV作成日"]),
                                    Memo = Utility.NulltoStr(dr["備考"]),
                                    UpDate = System.DateTime.Parse(Utility.NulltoStr(dr["更新年月日"])),
                                    Exception = Utility.StrtoInt(Utility.NulltoStr(dr["除外"]))
                                };
                            }
                        }
                    }
                }

                return registrationCard;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return registrationCard;
            }
        }

        /// <summary>
        /// 環境設定テーブルのデータを更新する：2026/08/18
        /// </summary>
        /// <param name="configData">TblConfigクラス</param>
        private void UpDate(TblConfig configData)
        {
            try
            {
                using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

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
                        com.Parameters.AddWithValue("@upDate", System.DateTime.Now.ToString());
                        com.Parameters.AddWithValue("@ID", configData.ID);
                        com.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 防犯カードテーブルのデータを更新する：2026/08/26
        /// </summary>
        /// <param name="workcardData">TblWorkcardクラス</param>
        private void UpDate(TblWorkcard workcardData)
        {
            try
            {
                using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    string sql = "UPDATE 防犯カード SET " +
                            "データ区分 = @datakbn, 画像名 = @imageName, 登録年 = @Year, 登録月 = @Month, 登録日 = @Day, " +
                            "登録番号 = @number, 車体番号 = @VehicleIdentificationNumber, メーカー = @Maker, 塗色 = @Color, " +
                            "車種 = @CarModel, 郵便番号1 = @ZipCode1, 郵便番号2 = @ZipCode2, 車両番号1 = @VehicleNumber1, " +
                            "車両番号2 = @VehicleNumber2, 車名 = @CarName, 住所漢字 = @AddressKanji," +
                            "住所1 = @Address1, 住所2 = @Address2, 氏名 = @Name, TEL携帯 = @Mobile1, TEL携帯2 = @Mobile2, TEL携帯3 = @Mobile3, " +
                            "PC名 = @PC, CSV作成日 = @CsvCreationDate, 備考 = @Memo, 更新年月日 = @UpDate, ラベル = @Label, " +
                            "処理担当者 = @Person, 確認 = @Confirmation " +
                            "WHERE ID = @ID";

                    using (SqlCommand com = new SqlCommand(sql, conn))
                    {
                        com.Parameters.AddWithValue("@datakbn", workcardData.DataCategory);
                        com.Parameters.AddWithValue("@imageName", workcardData.ImageFileName);
                        com.Parameters.AddWithValue("@year", workcardData.AddYear);
                        com.Parameters.AddWithValue("@month", workcardData.AddMonth);
                        com.Parameters.AddWithValue("@day", workcardData.AddDay);
                        com.Parameters.AddWithValue("@number", workcardData.Number);
                        com.Parameters.AddWithValue("@VehicleIdentificationNumber", workcardData.VehicleIdentificationNumber);
                        com.Parameters.AddWithValue("@Maker", workcardData.Maker);
                        com.Parameters.AddWithValue("@Color", workcardData.Color);
                        com.Parameters.AddWithValue("@CarModel", workcardData.CarModel);
                        com.Parameters.AddWithValue("@ZipCode1", workcardData.ZipCode1);
                        com.Parameters.AddWithValue("@ZipCode2", workcardData.ZipCode2);
                        com.Parameters.AddWithValue("@VehicleNumber1", workcardData.VehicleNumber1);
                        com.Parameters.AddWithValue("@VehicleNumber2", workcardData.VehicleNumber2);
                        com.Parameters.AddWithValue("@CarName", workcardData.CarName);
                        com.Parameters.AddWithValue("@AddressKanji", workcardData.AddressKanji);
                        com.Parameters.AddWithValue("@Address1", workcardData.Address1);
                        com.Parameters.AddWithValue("@Address2", workcardData.Address2);
                        com.Parameters.AddWithValue("@Name", workcardData.Name);
                        com.Parameters.AddWithValue("@Mobile1", workcardData.Mobile1);
                        com.Parameters.AddWithValue("@Mobile2", workcardData.Mobile2);
                        com.Parameters.AddWithValue("@Mobile3", workcardData.Mobile3);
                        com.Parameters.AddWithValue("@PC", workcardData.PC);
                        com.Parameters.AddWithValue("@CsvCreationDate", workcardData.CsvCreationDate);
                        com.Parameters.AddWithValue("@Memo", workcardData.Memo);
                        com.Parameters.AddWithValue("@UpDate", System.DateTime.Now.ToString());
                        com.Parameters.AddWithValue("@Label", workcardData.Label);
                        com.Parameters.AddWithValue("@Person", workcardData.Person);
                        com.Parameters.AddWithValue("@Confirmation", workcardData.Confirmation);
                        com.Parameters.AddWithValue("@ID", workcardData.ID);
                        com.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 防犯登録データテーブルを更新する：2026/08/31
        /// </summary>
        /// <param name="registrationCard">更新する防犯登録データ</param>
        private void UpDate(TblRegistrationCard registrationCard)
        {
            try
            {
                using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    string sql = "UPDATE 防犯登録データ SET " +
                            "データ区分 = @datakbn, 画像名 = @imageName, 登録年 = @Year, 登録月 = @Month, 登録日 = @Day, " +
                            "登録番号 = @number, 車体番号 = @VehicleIdentificationNumber, メーカー = @Maker, 塗色 = @Color, " +
                            "車種 = @CarModel, 郵便番号1 = @ZipCode1, 郵便番号2 = @ZipCode2, 車両番号1 = @VehicleNumber1, " +
                            "車両番号2 = @VehicleNumber2, 車名 = @CarName, 住所漢字 = @AddressKanji," +
                            "住所1 = @Address1, 住所2 = @Address2, 氏名 = @Name, TEL携帯 = @Mobile1, TEL携帯2 = @Mobile2, TEL携帯3 = @Mobile3, " +
                            "PC名 = @PC, CSV作成日 = @CsvCreationDate, 備考 = @Memo, 更新年月日 = @UpDate, 除外 = @Exception " +
                            "WHERE ID = @ID";

                    using (SqlCommand com = new SqlCommand(sql, conn))
                    {
                        com.Parameters.AddWithValue("@datakbn", registrationCard.DataCategory);
                        com.Parameters.AddWithValue("@imageName", registrationCard.ImageFileName);
                        com.Parameters.AddWithValue("@year", registrationCard.AddYear);
                        com.Parameters.AddWithValue("@month", registrationCard.AddMonth);
                        com.Parameters.AddWithValue("@day", registrationCard.AddDay);
                        com.Parameters.AddWithValue("@number", registrationCard.Number);
                        com.Parameters.AddWithValue("@VehicleIdentificationNumber", registrationCard.VehicleIdentificationNumber);
                        com.Parameters.AddWithValue("@Maker", registrationCard.Maker);
                        com.Parameters.AddWithValue("@Color", registrationCard.Color);
                        com.Parameters.AddWithValue("@CarModel", registrationCard.CarModel);
                        com.Parameters.AddWithValue("@ZipCode1", registrationCard.ZipCode1);
                        com.Parameters.AddWithValue("@ZipCode2", registrationCard.ZipCode2);
                        com.Parameters.AddWithValue("@VehicleNumber1", registrationCard.VehicleNumber1);
                        com.Parameters.AddWithValue("@VehicleNumber2", registrationCard.VehicleNumber2);
                        com.Parameters.AddWithValue("@CarName", registrationCard.CarName);
                        com.Parameters.AddWithValue("@AddressKanji", registrationCard.AddressKanji);
                        com.Parameters.AddWithValue("@Address1", registrationCard.Address1);
                        com.Parameters.AddWithValue("@Address2", registrationCard.Address2);
                        com.Parameters.AddWithValue("@Name", registrationCard.Name);
                        com.Parameters.AddWithValue("@Mobile1", registrationCard.Mobile1);
                        com.Parameters.AddWithValue("@Mobile2", registrationCard.Mobile2);
                        com.Parameters.AddWithValue("@Mobile3", registrationCard.Mobile3);
                        com.Parameters.AddWithValue("@PC", registrationCard.PC);
                        com.Parameters.AddWithValue("@CsvCreationDate", registrationCard.CsvCreationDate);
                        com.Parameters.AddWithValue("@Memo", registrationCard.Memo);
                        com.Parameters.AddWithValue("@UpDate", System.DateTime.Now.ToString());
                        com.Parameters.AddWithValue("@Exception", registrationCard.Exception);
                        com.Parameters.AddWithValue("@ID", registrationCard.ID);
                        com.ExecuteNonQuery();
                    }
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
        /// <param name="scandataList">挿入するTblScandataのリスト</param>
        public void Insert(List<TblScandata> scandataList)
        {
            using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
            {
                conn.Open();

                string sql = "INSERT INTO SCAN_DATA (データ区分, 画像名, 登録年, 登録月, 登録日, 登録番号," +
                    "車体番号, メーカー, 塗色, 車種, 郵便番号1, 郵便番号2, 車両番号1, 車両番号2, 車名, 住所漢字," +
                    "住所1, 住所2, 氏名, TEL携帯, TEL携帯2, TEL携帯3, PC名, CSV作成日, 備考, 更新年月日, ラベル, 処理担当者) " +
                    "values (@datakbn, @imageName, @Year, @Month, @Day, @number, @VehicleIdentificationNumber," +
                    "@Maker, @Color, @CarModel, @ZipCode1, @ZipCode2, @VehicleNumber1, @VehicleNumber2, @CarName, @AddressKanji," +
                    "@Address1, @Address2, @Name, @Mobile1, @Mobile2, @Mobile3, @PC, @CsvCreationDate, @Memo, @UpDate, @Label, @Person)";

                using (SqlCommand com = new SqlCommand(sql, conn))
                {
                    foreach (var scandata in scandataList)
                    {
                        com.Parameters.Clear();
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
                        com.Parameters.AddWithValue("@UpDate", System.DateTime.Now.ToString());
                        com.Parameters.AddWithValue("@Label", scandata.Label);
                        com.Parameters.AddWithValue("@Person", scandata.Person);
                        com.ExecuteNonQuery();
                    }
                }
            }
        }

        /// <summary>
        /// 防犯カードテーブルにデータを挿入する：2026/08/18
        /// </summary>
        /// <param name="scandataList">挿入するスキャンデータのリスト</param>
        /// <param name="pc">PC名</param>
        public void InsertWorkTbl(List<TblScandata> scandataList, string pc)
        {
            using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
            {
                conn.Open();

                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        // 取得したデータのラベルを取得する
                        string label = scandataList.FirstOrDefault()?.Label;

                        string sql = "INSERT INTO 防犯カード (データ区分, 画像名, 登録年, 登録月, 登録日, 登録番号," +
                            "車体番号, メーカー, 塗色, 車種, 郵便番号1, 郵便番号2, 車両番号1, 車両番号2, 車名, 住所漢字," +
                            "住所1, 住所2, 氏名, TEL携帯, TEL携帯2, TEL携帯3, PC名, CSV作成日, 備考, 更新年月日, ラベル, 処理担当者, 確認) " +
                            "values (@datakbn, @imageName, @Year, @Month, @Day, @number, @VehicleIdentificationNumber," +
                            "@Maker, @Color, @CarModel, @ZipCode1, @ZipCode2, @VehicleNumber1, @VehicleNumber2, @CarName, @AddressKanji," +
                            "@Address1, @Address2, @Name, @Mobile1, @Mobile2, @Mobile3, @PC, @CsvCreationDate, @Memo, @UpDate, @Label, @Person, @Confirmation)";

                        using (SqlCommand com = new SqlCommand(sql, conn, transaction))
                        {
                            foreach (var scandata in scandataList)
                            {
                                com.Parameters.Clear();
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
                                com.Parameters.AddWithValue("@VehicleNumber1", "");
                                com.Parameters.AddWithValue("@VehicleNumber2", "");
                                com.Parameters.AddWithValue("@CarName", "");
                                com.Parameters.AddWithValue("@AddressKanji", "");
                                com.Parameters.AddWithValue("@Address1", scandata.Address1);
                                com.Parameters.AddWithValue("@Address2", scandata.Address2);
                                com.Parameters.AddWithValue("@Name", scandata.Name);
                                com.Parameters.AddWithValue("@Mobile1", scandata.Mobile1);
                                com.Parameters.AddWithValue("@Mobile2", scandata.Mobile2);
                                com.Parameters.AddWithValue("@Mobile3", scandata.Mobile3);
                                com.Parameters.AddWithValue("@PC", pc);
                                com.Parameters.AddWithValue("@CsvCreationDate", scandata.CsvCreationDate);
                                com.Parameters.AddWithValue("@Memo", scandata.Memo);
                                com.Parameters.AddWithValue("@UpDate", System.DateTime.Now.ToString());
                                com.Parameters.AddWithValue("@Label", scandata.Label);
                                com.Parameters.AddWithValue("@Person", scandata.Person);
                                com.Parameters.AddWithValue("@Confirmation", 0);
                                com.ExecuteNonQuery();
                            }

                            // SCAN_DATAテーブルのデータを削除する
                            string deleteSql = "DELETE FROM SCAN_DATA WHERE ラベル = @Label";
                            using (SqlCommand deleteCom = new SqlCommand(deleteSql, conn, transaction))
                            {
                                deleteCom.Parameters.Add("@Label", SqlDbType.NVarChar).Value = label;
                                deleteCom.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        /// <summary>
        /// 防犯登録データテーブルにデータを挿入する：2026/08/28
        /// </summary>
        /// <param name="workcards">挿入する防犯カードデータのリスト</param>
        public void InsertRegistrationCardTbl(List<TblWorkcard> workcards)
        {
            using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
            {
                conn.Open();

                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        string sql = "INSERT INTO 防犯登録データ (データ区分, 画像名, 登録年, 登録月, 登録日, 登録番号," +
                            "車体番号, メーカー, 塗色, 車種, 郵便番号1, 郵便番号2, 車両番号1, 車両番号2, 車名, 住所漢字," +
                            "住所1, 住所2, 氏名, TEL携帯, TEL携帯2, TEL携帯3, PC名, CSV作成日, 備考, 更新年月日, 除外) " +
                            "values (@datakbn, @imageName, @Year, @Month, @Day, @number, @VehicleIdentificationNumber," +
                            "@Maker, @Color, @CarModel, @ZipCode1, @ZipCode2, @VehicleNumber1, @VehicleNumber2, @CarName, @AddressKanji," +
                            "@Address1, @Address2, @Name, @Mobile1, @Mobile2, @Mobile3, @PC, @CsvCreationDate, @Memo, @UpDate, @exception)";

                        using (SqlCommand com = new SqlCommand(sql, conn, transaction))
                        {
                            foreach (var workcard in workcards.OrderBy(w => w.ID))
                            {
                                com.Parameters.Clear();
                                com.Parameters.AddWithValue("@datakbn",         workcard.DataCategory);
                                com.Parameters.AddWithValue("@imageName",       workcard.ImageFileName);
                                com.Parameters.AddWithValue("@Year",            workcard.AddYear);
                                com.Parameters.AddWithValue("@Month",           workcard.AddMonth);
                                com.Parameters.AddWithValue("@Day",             workcard.AddDay);
                                com.Parameters.AddWithValue("@number",          workcard.Number);
                                com.Parameters.AddWithValue("@VehicleIdentificationNumber", workcard.VehicleIdentificationNumber);
                                com.Parameters.AddWithValue("@Maker",           workcard.Maker);
                                com.Parameters.AddWithValue("@Color",           workcard.Color);
                                com.Parameters.AddWithValue("@CarModel",        workcard.CarModel);
                                com.Parameters.AddWithValue("@ZipCode1",        workcard.ZipCode1);
                                com.Parameters.AddWithValue("@ZipCode2",        workcard.ZipCode2);
                                com.Parameters.AddWithValue("@VehicleNumber1",  "");
                                com.Parameters.AddWithValue("@VehicleNumber2",  "");
                                com.Parameters.AddWithValue("@CarName",         "");
                                com.Parameters.AddWithValue("@AddressKanji",    workcard.AddressKanji);
                                com.Parameters.AddWithValue("@Address1",        workcard.Address1);
                                com.Parameters.AddWithValue("@Address2",        workcard.Address2);
                                com.Parameters.AddWithValue("@Name",            workcard.Name);
                                com.Parameters.AddWithValue("@Mobile1",         workcard.Mobile1);
                                com.Parameters.AddWithValue("@Mobile2",         workcard.Mobile2);
                                com.Parameters.AddWithValue("@Mobile3",         workcard.Mobile3);
                                com.Parameters.AddWithValue("@PC",              workcard.PC);
                                com.Parameters.AddWithValue("@CsvCreationDate", workcard.CsvCreationDate);
                                com.Parameters.AddWithValue("@Memo",            workcard.Memo);
                                com.Parameters.AddWithValue("@UpDate",          System.DateTime.Now.ToString());
                                com.Parameters.AddWithValue("@exception",       global.flgOff);
                                com.ExecuteNonQuery();

                                // 防犯カードデータを削除する
                                string deleteSql = "DELETE FROM 防犯カード WHERE ID = @ID";
                                using (SqlCommand deleteCom = new SqlCommand(deleteSql, conn, transaction))
                                {
                                    deleteCom.Parameters.AddWithValue("@ID", workcard.ID);
                                    deleteCom.ExecuteNonQuery();
                                }

                                // 画像ファイルを移動する
                                if (System.IO.File.Exists(Properties.Settings.Default.scanDataPath + workcard.ImageFileName))
                                {
                                    // 既に画像ファイルが存在する場合は削除する
                                    if (System.IO.File.Exists(Properties.Settings.Default.imgPath + workcard.ImageFileName))
                                    {
                                        System.IO.File.Delete(Properties.Settings.Default.imgPath + workcard.ImageFileName);
                                    }

                                    // 画像ファイルを移動する
                                    System.IO.File.Copy(Properties.Settings.Default.scanDataPath + workcard.ImageFileName, Properties.Settings.Default.imgPath + workcard.ImageFileName);
                                    System.IO.File.Delete(Properties.Settings.Default.scanDataPath + workcard.ImageFileName);
                                }
                            }
                        }

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
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
                com.Parameters.AddWithValue("@UpDate", System.DateTime.Now.ToString());
                com.Parameters.AddWithValue("@Label", scandata.Label);
                com.Parameters.AddWithValue("@Person", scandata.Person);
                com.ExecuteNonQuery();
            }
        }

        public List<T> Read<T>(ScandataParameter param)
        {
            // SCANDATAのとき
            if (typeof(T) == typeof(TblScandata))
            {
                return Read(param) as List<T>;
            }

            // 防犯登録カードのとき
            if (typeof(T) == typeof(TblRegistrationCard))
            {
                return ReadRegistrationCard(param) as List<T>;
            }

            MessageBox.Show("Invalid Data Class");
            return null;
        }

        public List<T> Read<T>()
        {
            // CSV作成履歴のとき
            if (typeof(T) == typeof(TblCSVCreationHistory))
            {
                return ReadCSVCreationHistory() as List<T>;
            }

            MessageBox.Show("Invalid Data Class");
            return null;
        }

        public List<ClsLabelCount> LabelCount()
        {
            var LabelCounts = new List<ClsLabelCount>();

            try
            {
                using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();
                    string sql = "SELECT ラベル, 処理担当者, COUNT(*) AS 件数 FROM SCAN_DATA GROUP BY ラベル, 処理担当者 ORDER BY ラベル";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                ClsLabelCount labelCount = new ClsLabelCount
                                {
                                    Label = Utility.NulltoStr(dr["ラベル"]),
                                    Person = Utility.NulltoStr(dr["処理担当者"]),
                                    Count = Utility.NulltoStr(dr["件数"])
                                };
                                LabelCounts.Add(labelCount);
                            }
                        }
                    }
                }
                return LabelCounts;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return LabelCounts;
            }
        }

        /// <summary>
        /// SCAN_DATAテーブルのデータを取得する：2026/08/28
        /// </summary>
        /// <param name="param">検索パラメータ</param>
        /// <returns>SCAN_DATAのリスト</returns>
        public List<TblScandata> Read(ScandataParameter param)
        {
            var lines = new List<TblScandata>();

            try
            {
                using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    string sql = "SELECT ID,データ区分,画像名,登録年,登録月,登録日,登録番号,車体番号,メーカー,塗色,車種,郵便番号1,郵便番号2," +
                        "住所1,氏名,TEL携帯,TEL携帯2,TEL携帯3,ラベル,処理担当者,PC名,CSV作成日,備考,更新年月日 FROM SCAN_DATA " +
                        "WHERE (" +
                        "(@DataCategory IS NULL OR データ区分 = @DataCategory) AND " +
                        "(@AddYear IS NULL OR 登録年 = @AddYear) AND " +
                        "(@AddMonth IS NULL OR 登録月 = @AddMonth OR 登録月 = @AddMonth2) AND " +
                        "(@AddDay IS NULL OR 登録日 = @AddDay OR 登録日 = @AddDay2) AND " +
                        "(@Number IS NULL OR 登録番号 LIKE '%' + @Number + '%') AND " +
                        "(@VehicleIdentificationNumber IS NULL OR 車体番号 LIKE '%' + @VehicleIdentificationNumber + '%') AND " +
                        "(@Maker IS NULL OR メーカー LIKE '%' + @Maker + '%') AND " +
                        "(@Color IS NULL OR 塗色 LIKE '%' + @Color + '%') AND " +
                        "(@CarModel IS NULL OR 車種 = @CarModel) AND " +
                        "(@ZipCode1 IS NULL OR 郵便番号1 LIKE '%' + @ZipCode1 + '%') AND " +
                        "(@ZipCode2 IS NULL OR 郵便番号2 LIKE '%' + @ZipCode2 + '%') AND " +
                        "(@Address1 IS NULL OR 住所1 LIKE '%' + @Address1 + '%') AND " +
                        "(@Name IS NULL OR 氏名 LIKE '%' + @Name + '%') AND " +
                        "(@Mobile1 IS NULL OR TEL携帯 LIKE '%' + @Mobile1 + '%') AND " +
                        "(@Mobile2 IS NULL OR TEL携帯2 LIKE '%' + @Mobile2 + '%') AND " +
                        "(@Mobile3 IS NULL OR TEL携帯3 LIKE '%' + @Mobile3 + '%') AND " +
                        "(@Label IS NULL OR ラベル LIKE '%' + @Label + '%') AND " +
                        "(@Person IS NULL OR 処理担当者 LIKE '%' + @Person + '%')) " +
                        "ORDER BY 登録番号 OPTION (RECOMPILE)";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.CommandTimeout = 120; // 秒。デフォルト30から一時的に伸ばして様子を見る
                        var dataCategoryParameter = cmd.Parameters.Add("@DataCategory", System.Data.SqlDbType.Int);
                        dataCategoryParameter.Value = param.DataCategory.HasValue ? (object)param.DataCategory.Value : DBNull.Value;

                        cmd.Parameters.Add("@AddYear", SqlDbType.NVarChar, 2).Value = ToParam(param.AddYear);

                        cmd.Parameters.Add("@AddMonth", SqlDbType.NVarChar, 2).Value = ToParam(param.AddMonth);
                        if (!string.IsNullOrEmpty(param.AddMonth))
                        {
                            cmd.Parameters.Add("@AddMonth2", SqlDbType.NVarChar, 2).Value = param.AddMonth.PadLeft(2, '0');
                        }
                        else
                        {
                            cmd.Parameters.Add("@AddMonth2", SqlDbType.NVarChar, 2).Value = DBNull.Value;
                        }

                        cmd.Parameters.Add("@AddDay", SqlDbType.NVarChar, 2).Value = ToParam(param.AddDay);
                        if (!string.IsNullOrEmpty(param.AddDay))
                        {
                            cmd.Parameters.Add("@AddDay2", SqlDbType.NVarChar, 2).Value = param.AddDay.PadLeft(2, '0');
                        }
                        else
                        {
                            cmd.Parameters.Add("@AddDay2", SqlDbType.NVarChar, 2).Value = DBNull.Value;
                        }

                        cmd.Parameters.Add("@Number", SqlDbType.NVarChar, 20).Value = ToParam(param.Number);
                        cmd.Parameters.Add("@VehicleIdentificationNumber", SqlDbType.NVarChar, 20).Value = ToParam(param.VehicleIdentificationNumber);
                        cmd.Parameters.Add("@Maker", SqlDbType.NVarChar, 10).Value = ToParam(param.Maker);
                        cmd.Parameters.Add("@Color", SqlDbType.NVarChar, 6).Value = ToParam(param.Color);

                        var CarModelParameter = cmd.Parameters.Add("@CarModel", System.Data.SqlDbType.Int);
                        CarModelParameter.Value = param.CarModel.HasValue ? (object)param.CarModel.Value : DBNull.Value;

                        cmd.Parameters.Add("@ZipCode1", SqlDbType.NVarChar, 3).Value = ToParam(param.ZipCode1);
                        cmd.Parameters.Add("@ZipCode2", SqlDbType.NVarChar, 4).Value = ToParam(param.ZipCode2);
                        cmd.Parameters.Add("@Address1", SqlDbType.NVarChar, 40).Value = ToParam(param.Address1);
                        cmd.Parameters.Add("@Name", SqlDbType.NVarChar, 16).Value = ToParam(param.Name);
                        cmd.Parameters.Add("@Mobile1", SqlDbType.NVarChar, 4).Value = ToParam(param.Mobile1);
                        cmd.Parameters.Add("@Mobile2", SqlDbType.NVarChar, 4).Value = ToParam(param.Mobile2);
                        cmd.Parameters.Add("@Mobile3", SqlDbType.NVarChar, 4).Value = ToParam(param.Mobile3);
                        cmd.Parameters.Add("@Label", SqlDbType.NVarChar, 255).Value = ToParam(param.Label);
                        cmd.Parameters.Add("@Person", SqlDbType.NVarChar, 255).Value = ToParam(param.Person);

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
                                    VehicleNumber1 = string.Empty,
                                    VehicleNumber2 = string.Empty,
                                    CarName = string.Empty,
                                    Address1 = Utility.NulltoStr(dr["住所1"]),
                                    Address2 = string.Empty,
                                    Name = Utility.NulltoStr(dr["氏名"]),
                                    Mobile1 = Utility.NulltoStr(dr["TEL携帯"]),
                                    Mobile2 = Utility.NulltoStr(dr["TEL携帯2"]),
                                    Mobile3 = Utility.NulltoStr(dr["TEL携帯3"]),
                                    PC = Utility.NulltoStr(dr["PC名"]),
                                    CsvCreationDate = Utility.NulltoStr(dr["CSV作成日"]),
                                    Memo = Utility.NulltoStr(dr["備考"]),
                                    UpDate = System.DateTime.Parse(Utility.NulltoStr(dr["更新年月日"])),
                                    Label = Utility.NulltoStr(dr["ラベル"]),
                                    Person = Utility.NulltoStr(dr["処理担当者"])
                                };

                                lines.Add(scandata);
                            }
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

        /// <summary>
        /// 防犯登録カードテーブルのデータを取得する：2026/08/28
        /// </summary>
        /// <param name="param">検索パラメータ</param>
        /// <returns>防犯登録カードのリスト</returns>
        public List<TblRegistrationCard> ReadRegistrationCard(ScandataParameter param)
        {
            var lines = new List<TblRegistrationCard>();

            try
            {
                using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    string sql = "SELECT ID,データ区分,画像名,登録年,登録月,登録日,登録番号,車体番号,メーカー,塗色,車種,郵便番号1,郵便番号2," +
                        "住所漢字, 住所1,氏名,TEL携帯,TEL携帯2,TEL携帯3,PC名,CSV作成日,備考,更新年月日,除外 FROM 防犯登録データ " +
                        "WHERE (" +
                        "(@DataCategory IS NULL OR データ区分 = @DataCategory) AND " +
                        "(@AddYear IS NULL OR 登録年 = @AddYear) AND " +
                        "(@AddMonth IS NULL OR 登録月 = @AddMonth OR 登録月 = @AddMonth2) AND " +
                        "(@AddDay IS NULL OR 登録日 = @AddDay OR 登録日 = @AddDay2) AND " +
                        "(@Number IS NULL OR 登録番号 LIKE '%' + @Number + '%') AND " +
                        "(@VehicleIdentificationNumber IS NULL OR 車体番号 LIKE '%' + @VehicleIdentificationNumber + '%') AND " +
                        "(@Maker IS NULL OR メーカー LIKE '%' + @Maker + '%') AND " +
                        "(@Color IS NULL OR 塗色 LIKE '%' + @Color + '%') AND " +
                        "(@CarModel IS NULL OR 車種 = @CarModel) AND " +
                        "(@ZipCode1 IS NULL OR 郵便番号1 LIKE '%' + @ZipCode1 + '%') AND " +
                        "(@ZipCode2 IS NULL OR 郵便番号2 LIKE '%' + @ZipCode2 + '%') AND " +
                        "(@Address1 IS NULL OR 住所1 LIKE '%' + @Address1 + '%') AND " +
                        "(@Name IS NULL OR 氏名 LIKE '%' + @Name + '%') AND " +
                        "(@Mobile1 IS NULL OR TEL携帯 LIKE '%' + @Mobile1 + '%') AND " +
                        "(@Mobile2 IS NULL OR TEL携帯2 LIKE '%' + @Mobile2 + '%') AND " +
                        "(@Mobile3 IS NULL OR TEL携帯3 LIKE '%' + @Mobile3 + '%') AND " +
                        "(@Exception IS NULL OR 除外 = @Exception)) ";

                    // 県警用CSV作成済み条件を追加
                    if (param.CsvCreation != null)
                    {
                        if (param.CsvCreation == 0)
                        {
                            // CSV作成済み
                            if (string.IsNullOrEmpty(param.CsvCreationDate))
                            {
                                // 作成日付指定なし
                                sql += "AND (CSV作成日 IS NOT NULL AND CSV作成日 != '') ";
                            }
                            else
                            {
                                // 作成日付指定
                                sql += "AND (CSV作成日 IS NOT NULL AND CSV作成日 != '' AND CSV作成日 LIKE '%' + @CsvCreationDate + '%') ";
                            }
                        }
                        else if (param.CsvCreation == 1)
                        {
                            // CSV未作成
                            sql += "AND (CSV作成日 IS NULL OR CSV作成日 = '') ";
                        }
                    }

                    sql += "ORDER BY 登録番号 OPTION (RECOMPILE)";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.CommandTimeout = 120; // 秒。デフォルト30から一時的に伸ばして様子を見る

                        var dataCategoryParameter = cmd.Parameters.Add("@DataCategory", System.Data.SqlDbType.Int);
                        dataCategoryParameter.Value = param.DataCategory.HasValue ? (object)param.DataCategory.Value : DBNull.Value;

                        cmd.Parameters.Add("@AddYear", SqlDbType.NVarChar, 2).Value = ToParam(param.AddYear);

                        cmd.Parameters.Add("@AddMonth", SqlDbType.NVarChar, 2).Value = ToParam(param.AddMonth);
                        if (!string.IsNullOrEmpty(param.AddMonth))
                        {
                            cmd.Parameters.Add("@AddMonth2", SqlDbType.NVarChar, 2).Value = param.AddMonth.PadLeft(2, '0');
                        }
                        else
                        {
                            cmd.Parameters.Add("@AddMonth2", SqlDbType.NVarChar, 2).Value = DBNull.Value;
                        }

                        cmd.Parameters.Add("@AddDay", SqlDbType.NVarChar, 2).Value = ToParam(param.AddDay);
                        if (!string.IsNullOrEmpty(param.AddDay))
                        {
                            cmd.Parameters.Add("@AddDay2", SqlDbType.NVarChar, 2).Value = param.AddDay.PadLeft(2, '0');
                        }
                        else
                        {
                            cmd.Parameters.Add("@AddDay2", SqlDbType.NVarChar, 2).Value = DBNull.Value;
                        }

                        cmd.Parameters.Add("@Number", SqlDbType.NVarChar, 20).Value = ToParam(param.Number);
                        cmd.Parameters.Add("@VehicleIdentificationNumber", SqlDbType.NVarChar, 20).Value = ToParam(param.VehicleIdentificationNumber);
                        cmd.Parameters.Add("@Maker", SqlDbType.NVarChar, 10).Value = ToParam(param.Maker);
                        cmd.Parameters.Add("@Color", SqlDbType.NVarChar, 6).Value = ToParam(param.Color);

                        var CarModelParameter = cmd.Parameters.Add("@CarModel", System.Data.SqlDbType.Int);
                        CarModelParameter.Value = param.CarModel.HasValue ? (object)param.CarModel.Value : DBNull.Value;

                        cmd.Parameters.Add("@ZipCode1", SqlDbType.NVarChar, 3).Value = ToParam(param.ZipCode1);
                        cmd.Parameters.Add("@ZipCode2", SqlDbType.NVarChar, 4).Value = ToParam(param.ZipCode2);
                        cmd.Parameters.Add("@Address1", SqlDbType.NVarChar, 40).Value = ToParam(param.Address1);
                        cmd.Parameters.Add("@Name", SqlDbType.NVarChar, 16).Value = ToParam(param.Name);
                        cmd.Parameters.Add("@Mobile1", SqlDbType.NVarChar, 4).Value = ToParam(param.Mobile1);
                        cmd.Parameters.Add("@Mobile2", SqlDbType.NVarChar, 4).Value = ToParam(param.Mobile2);
                        cmd.Parameters.Add("@Mobile3", SqlDbType.NVarChar, 4).Value = ToParam(param.Mobile3);
                        cmd.Parameters.Add("@CsvCreationDate", SqlDbType.NVarChar, 255).Value = ToParam(param.CsvCreationDate);

                        var ExceptionParameter = cmd.Parameters.Add("@Exception", System.Data.SqlDbType.Int);
                        ExceptionParameter.Value = param.Exception.HasValue ? (object)param.Exception.Value : DBNull.Value;

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                var registrationCard = new TblRegistrationCard
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
                                    VehicleNumber1 = string.Empty,
                                    VehicleNumber2 = string.Empty,
                                    CarName = string.Empty,
                                    AddressKanji = Utility.NulltoStr(dr["住所漢字"]),
                                    Address1 = Utility.NulltoStr(dr["住所1"]),
                                    Address2 = string.Empty,
                                    Name = Utility.NulltoStr(dr["氏名"]),
                                    Mobile1 = Utility.NulltoStr(dr["TEL携帯"]),
                                    Mobile2 = Utility.NulltoStr(dr["TEL携帯2"]),
                                    Mobile3 = Utility.NulltoStr(dr["TEL携帯3"]),
                                    PC = Utility.NulltoStr(dr["PC名"]),
                                    CsvCreationDate = Utility.NulltoStr(dr["CSV作成日"]),
                                    Memo = Utility.NulltoStr(dr["備考"]),
                                    UpDate = System.DateTime.Parse(Utility.NulltoStr(dr["更新年月日"])),
                                    Exception = Utility.StrtoInt(Utility.NulltoStr(dr["除外"]))
                                };

                                lines.Add(registrationCard);
                            }
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

        /// <summary>
        /// 指定されたラベルに一致するSCAN_DATAテーブルのデータを取得する
        /// </summary>
        /// <param name="label">ラベル</param>
        /// <returns>指定されたラベルに一致するSCAN_DATAテーブルのデータのリスト</returns>
        public List<TblScandata> ReadLabel(string label)
        {
            var lines = new List<TblScandata>();

            try
            {
                using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    string sql = "SELECT ID,データ区分,画像名,登録年,登録月,登録日,登録番号,車体番号,メーカー,塗色,車種,郵便番号1,郵便番号2," +
                        "住所1,氏名,TEL携帯,TEL携帯2,TEL携帯3,ラベル,処理担当者,PC名,CSV作成日,備考,更新年月日 FROM SCAN_DATA " +
                        "WHERE (ラベル = @Label) ";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.CommandTimeout = 120; // 秒。デフォルト30から一時的に伸ばして様子を見る                                               
                        cmd.Parameters.Add("@Label", SqlDbType.NVarChar, 255).Value = label;

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
                                    VehicleNumber1 = string.Empty,
                                    VehicleNumber2 = string.Empty,
                                    CarName = string.Empty,
                                    Address1 = Utility.NulltoStr(dr["住所1"]),
                                    Address2 = string.Empty,
                                    Name = Utility.NulltoStr(dr["氏名"]),
                                    Mobile1 = Utility.NulltoStr(dr["TEL携帯"]),
                                    Mobile2 = Utility.NulltoStr(dr["TEL携帯2"]),
                                    Mobile3 = Utility.NulltoStr(dr["TEL携帯3"]),
                                    PC = Utility.NulltoStr(dr["PC名"]),
                                    CsvCreationDate = Utility.NulltoStr(dr["CSV作成日"]),
                                    Memo = Utility.NulltoStr(dr["備考"]),
                                    UpDate = System.DateTime.Parse(Utility.NulltoStr(dr["更新年月日"])),
                                    Label = Utility.NulltoStr(dr["ラベル"]),
                                    Person = Utility.NulltoStr(dr["処理担当者"])
                                };

                                lines.Add(scandata);
                            }
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


        public List<TblWorkcard> ReadPc(string pc)
        {
            var lines = new List<TblWorkcard>();

            try
            {
                using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    string sql = "SELECT ID,データ区分,画像名,登録年,登録月,登録日,登録番号,車体番号,メーカー,塗色,車種,郵便番号1,郵便番号2," +
                        "住所漢字,住所1,氏名,TEL携帯,TEL携帯2,TEL携帯3,ラベル,処理担当者,PC名,CSV作成日,備考,更新年月日 FROM 防犯カード " +
                        "WHERE (PC名 = @PC) ";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.CommandTimeout = 120; // 秒。デフォルト30から一時的に伸ばして様子を見る                                               
                        cmd.Parameters.Add("@PC", SqlDbType.NVarChar, 255).Value = pc;

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                TblWorkcard work = new TblWorkcard
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
                                    VehicleNumber1 = string.Empty,
                                    VehicleNumber2 = string.Empty,
                                    CarName = string.Empty,
                                    AddressKanji = Utility.NulltoStr(dr["住所漢字"]),
                                    Address1 = Utility.NulltoStr(dr["住所1"]),
                                    Address2 = string.Empty,
                                    Name = Utility.NulltoStr(dr["氏名"]),
                                    Mobile1 = Utility.NulltoStr(dr["TEL携帯"]),
                                    Mobile2 = Utility.NulltoStr(dr["TEL携帯2"]),
                                    Mobile3 = Utility.NulltoStr(dr["TEL携帯3"]),
                                    PC = Utility.NulltoStr(dr["PC名"]),
                                    CsvCreationDate = Utility.NulltoStr(dr["CSV作成日"]),
                                    Memo = Utility.NulltoStr(dr["備考"]),
                                    UpDate = System.DateTime.Parse(Utility.NulltoStr(dr["更新年月日"])),
                                    Label = Utility.NulltoStr(dr["ラベル"]),
                                    Person = Utility.NulltoStr(dr["処理担当者"])
                                };

                                lines.Add(work);
                            }
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

        /// <summary>
        /// CSV作成履歴テーブルのデータを取得する
        /// </summary>
        /// <returns>CSV作成履歴のリスト</returns>
        public List<TblCSVCreationHistory> ReadCSVCreationHistory()
        {
            var lines = new List<TblCSVCreationHistory>();

            try
            {
                using (var conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    string sql = "SELECT ID, 作成年月日, 開始年月日, 終了年月日, 出力件数, PC名, 摘要 FROM CSV作成履歴 ";

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.CommandTimeout = 120; // 秒。デフォルト30から一時的に伸ばして様子を見る    

                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                TblCSVCreationHistory history = new TblCSVCreationHistory
                                {
                                    ID = Utility.StrtoInt(dr["ID"].ToString()),
                                    CreationDate = System.DateTime.Parse(Utility.NulltoStr(dr["作成年月日"])),
                                    Outputs = Utility.StrtoInt(Utility.NulltoStr(dr["出力件数"])),
                                    PC = Utility.NulltoStr(dr["PC名"]),
                                    Memo = Utility.NulltoStr(dr["摘要"])
                                };

                                lines.Add(history);
                            }
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

        /// <summary>
        /// 指定された文字列をSQLパラメータに変換する。空文字列の場合はDBNull.Valueを返す。
        /// </summary>
        /// <param name="value">変換する文字列</param>
        /// <returns>SQLパラメータとして使用できるオブジェクト</returns>
        static object ToParam(string value)
        {
            return string.IsNullOrEmpty(value) ? (object)DBNull.Value : value;
        }

        /// <summary>
        /// 指定されたSQLクエリを実行し、結果の件数を取得する
        /// </summary>
        /// <param name="sql">実行するSQLクエリ</param>
        /// <returns>結果の件数</returns>
        private int GetCount(string sql)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        var rtn = cmd.ExecuteScalar();
                        return Convert.ToInt32(rtn);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return 0;
            }
        }

        /// <summary>
        /// 指定されたSQLクエリを実行し、指定されたPC名に一致する結果の件数を取得する
        /// </summary>
        /// <param name="sql">実行するSQLクエリ</param>
        /// <param name="pc">PC名</param>
        /// <returns>結果の件数</returns>
        private int GetCount(string sql, string pc)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@val", pc);
                        var rtn = cmd.ExecuteScalar();
                        return Convert.ToInt32(rtn);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return 0;
            }
        }

        /// <summary>
        /// 指定されたSQLクエリを実行し、指定された登録番号とPC名に一致する結果の件数を取得する
        /// </summary>
        /// <param name="sql">実行するSQLクエリ</param>
        /// <param name="number">登録番号</param>
        /// <param name="pc">PC名</param>
        /// <returns>結果の件数</returns>
        private int GetCount(string sql, string number, string pc)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@Number", number);
                        cmd.Parameters.AddWithValue("@val", pc);
                        var rtn = cmd.ExecuteScalar();
                        return Convert.ToInt32(rtn);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return 0;
            }
        }

        public bool Delete<T>(string id)
        {
            // 防犯カードのとき
            if (typeof(T) == typeof(TblWorkcard))
            {
                return DeleteWorkCard(id);
            }

            MessageBox.Show("Invalid Data Class");
            return false;
        }

        /// <summary>
        /// 指定されたIDに一致する防犯カードテーブルのデータを削除する
        /// </summary>
        /// <param name="id">削除する防犯カードのID</param>
        /// <returns>削除が成功した場合はtrue、それ以外の場合はfalse</returns>
        private bool DeleteWorkCard(string id)
        {
            SqlTransaction tran = null;

            try
            {
                using (SqlConnection conn = new SqlConnection(sqlBuilder.ConnectionString))
                {
                    conn.Open();

                    // 防犯カードデータ
                    string sql = "DELETE FROM 防犯カード WHERE ID = @ID";

                    using (SqlCommand cmd = new SqlCommand(sql, conn, tran))
                    {
                        cmd.Parameters.AddWithValue("@ID", id);
                        cmd.ExecuteNonQuery();
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

    }
}
