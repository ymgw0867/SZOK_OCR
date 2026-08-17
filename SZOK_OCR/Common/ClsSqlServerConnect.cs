using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SZOK_OCR.Common
{
    public class ClsSqlServerConnect
    {
        protected SqlConnectionStringBuilder sqlBuilder { get; }

        public ClsSqlServerConnect(string sServerName, string sLogin, string sPass, string sDatabase)
        {
            try
            {
                // データベース接続文字列を作成
                SqlConnectionStringBuilder Builder = new SqlConnectionStringBuilder
                {
                    DataSource = sServerName,
                    UserID = sLogin,
                    Password = sPass,
                    InitialCatalog = sDatabase
                };

                sqlBuilder = Builder;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        protected SqlConnection GetSQLConnection()
        {
            return new SqlConnection(sqlBuilder.ConnectionString);
        }
    }
}
