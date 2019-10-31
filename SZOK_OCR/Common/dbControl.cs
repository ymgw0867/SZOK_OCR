using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;
using JS_OCR.Common;

namespace JS_OCR.Common
{
    class dbControl
    {
        /// <summary>
        /// DataControlクラスの基本クラス
        /// </summary>
        public class BaseControl
        {
            private DBConnect DBConnect;
            protected SqlConnection dbControlCn;

            // BaseControlのコンストラクタ。DBConnectクラスのインスタンスを作成します。
            public BaseControl(string dbName)
            {
                // データベースをオープンする
                DBConnect = new DBConnect(dbName);
            }

            // データベースに接続しコネクション情報を返す
            public SqlConnection GetConnection()
            {
                dbControlCn = DBConnect.Cn;
                return DBConnect.Cn;
            }
        }

        public class DataControl : BaseControl
        {
            // データコントロールクラスのコンストラクタ
            public DataControl(string dbName):base(dbName)
            {
            }

            /// <summary>
            /// データベース接続解除
            /// </summary>
            public void Close()
            {
                if (dbControlCn.State == ConnectionState.Open)
                {
                    dbControlCn.Close();
                }
            }

            /// <summary>
            /// 任意のSQLを実行する
            /// </summary>
            /// <param name="tempSql">SQL文</param>
            /// <returns>成功 : true, 失敗 : false</returns>
            public bool FreeSql(string tempSql)
            {
                bool rValue = false;

                try
                {
                    SqlCommand sCom = new SqlCommand();
                    sCom.CommandText = tempSql;
                    sCom.Connection = GetConnection();

                    //SQLの実行
                    sCom.ExecuteNonQuery();
                    rValue = true;
                }
                catch (Exception ex)
                {
                    rValue = false;
                }

                return rValue;
            }

            /// <summary>
            /// データリーダーを取得する
            /// </summary>
            /// <param name="tempSQL">SQL文</param>
            /// <returns>データリーダー</returns>
            public SqlDataReader FreeReader(string tempSQL)
            {
                SqlCommand sCom = new SqlCommand();
                sCom.CommandText = tempSQL;
                sCom.Connection = GetConnection();
                SqlDataReader dR = sCom.ExecuteReader();

                return dR;
            }

            /// <summary>
            /// データリーダーを取得する
            /// </summary>
            /// <param name="tempSQL">SQL文</param>
            /// <returns>データリーダー</returns>
            public SqlDataReader sqlReader(string tempSQL)
            {
                SqlCommand sCom = new SqlCommand();
                sCom.CommandText = tempSQL;
                sCom.Connection = GetConnection();
                SqlDataReader dR = sCom.ExecuteReader();

                return dR;
            }

            /// -----------------------------------------------------------
            /// <summary>
            ///     社員番号を指定して社員情報を取得します </summary>
            /// <param name="sYY">
            ///     基準年</param>
            /// <param name="sMM">
            ///     基準月</param>
            /// <returns>
            ///     データリーダー</returns>
            /// -----------------------------------------------------------
            public SqlDataReader GetEmployeeBase(string sNo)
            {
                // SQLServer接続
                SqlDataReader dRs;
                StringBuilder sb = new StringBuilder();

                sb.Append("select CODE,NAME from SHAIN1 ");
                sb.Append("where DEL = 1 and TAIKYU = 0 and CODE = '" + sNo + "'");
                dRs = FreeReader(sb.ToString());
                return dRs;
            }

            /// -----------------------------------------------------------
            /// <summary>
            ///     社員情報を取得します </summary>
            /// <param name="sYY">
            ///     基準年</param>
            /// <param name="sMM">
            ///     基準月</param>
            /// <returns>
            ///     データリーダー</returns>
            /// -----------------------------------------------------------
            public SqlDataReader GetEmployeeBase()
            {
                // SQLServer接続
                SqlDataReader dRs;
                StringBuilder sb = new StringBuilder();

                sb.Append("select DEL,INCODE,CODE,NAME,TAIKYU from SHAIN1 ");
                sb.Append("where DEL = 1 order by INCODE");
                dRs = FreeReader(sb.ToString());
                return dRs;
            }
        }

        /// <summary>
        /// SQLServerデータベース接続クラス
        /// </summary>
        public class DBConnect
        {
            SqlConnection cn = new SqlConnection();

            public SqlConnection Cn
            {
                get
                {
                    return cn;
                }
            }

            public DBConnect(string dbName)
            {
                try
                {
                    // データベース接続文字列
                    cn.ConnectionString = "Data Source=" + Properties.Settings.Default.SQLServerName + ";Initial Catalog=" + dbName + ";Integrated Security=True";
                    //cn.ConnectionString = Properties.Settings.Default.sqlConnectionStr;
                    cn.Open();
                }

                catch (Exception e)
                {
                    throw e;
                }
            }
        }
    }
}
