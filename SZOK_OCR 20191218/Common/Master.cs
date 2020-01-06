using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace JS_OCR.Common
{
    public class Master
    {
        public SqlCommand sCom = new SqlCommand();
        protected StringBuilder sb = new StringBuilder();

        public Master()
        {
        }

        public void dbConnect()
        {
            // データベース接続文字列
            SqlConnection Cn = new SqlConnection();
            sb.Clear();
            sb.Append("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=");
            sb.Append(Properties.Settings.Default.mdbPath);
            sb.Append(global.MDBFILE);
            Cn.ConnectionString = sb.ToString();
            Cn.Open();

            sCom.Connection = Cn;
        }
    }

    public class msConfig : Master
    {
        /// <summary>
        /// 環境設定マスター新規登録
        /// </summary>
        /// <param name="sID">キー</param>
        /// <param name="sSYEAR">年</param>
        /// <param name="sSMONTH">月</param>
        /// <param name="sPath">受け渡しデータ作成先パス</param>
        public void Insert(int sID, string sSYEAR, string sSMONTH, string sPath, string sArchived)
        {
            try
            {
                sb.Clear();
                sb.Append("insert into 環境設定 (");
                sb.Append("ID,年,月,受け渡しデータ作成パス,データ保存月数,更新年月日) values (");
                sb.Append("?,?,?,?,?,?)");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", sID);
                sCom.Parameters.AddWithValue("@year", sSYEAR);
                sCom.Parameters.AddWithValue("@Month", sSMONTH);
                sCom.Parameters.AddWithValue("@Path", sPath);
                sCom.Parameters.AddWithValue("@Arc", sArchived);
                sCom.Parameters.AddWithValue("@update", DateTime.Today.ToShortDateString());
                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }            
        }

        /// <summary>
        /// 環境設定マスター更新
        /// </summary>
        /// <param name="sSYEAR">年</param>
        /// <param name="sSMONTH">月</param>
        /// <param name="sPath">受け渡しデータ作成先パス</param>
        public void UpDate(string sSYEAR, string sSMONTH, string sPath, string sArchived)
        {
            try
            {
                sb.Clear();
                sb.Append("update 環境設定 set ");
                sb.Append("年=?,月=?,受け渡しデータ作成パス=?,データ保存月数=?,更新年月日=?");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@year", sSYEAR);
                sCom.Parameters.AddWithValue("@Month", sSMONTH);
                sCom.Parameters.AddWithValue("@path", sPath);
                sCom.Parameters.AddWithValue("@arc", sArchived);
                sCom.Parameters.AddWithValue("@update", DateTime.Today.ToShortDateString());

                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }
        
        /// <summary>
        /// 環境設定マスター取得
        /// </summary>
        /// <param name="sID">環境設定キー</param>
        public SqlDataReader Select(int sID)
        {
            SqlDataReader dr = null;

            try
            {
                sb.Clear();
                sb.Append("select * from 環境設定 ");
                sb.Append("where ID = ?");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", sID);
                dr = sCom.ExecuteReader();
                return dr;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return dr;
            }
            finally
            {
                //if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// -------------------------------------------------
        /// <summary>
        ///     環境設定データを取得する </summary>
        /// -------------------------------------------------
        public void GetCommonYearMonth()
        {
            JSDataSetTableAdapters.環境設定TableAdapter adp = new JSDataSetTableAdapters.環境設定TableAdapter();
            JSDataSet.環境設定DataTable cTbl = new JSDataSet.環境設定DataTable(); 

            try
            {
                adp.Fill(cTbl);
                JSDataSet.環境設定Row r = cTbl.FindByID(global.configKEY);

                global.cnfYear = r.年;
                global.cnfMonth = r.月;
                global.cnfPath = r.受け渡しデータ作成パス;
                global.cnfArchived = r.データ保存月数;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定年月取得", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
            }
        }
    }
}
