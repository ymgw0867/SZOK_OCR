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
        /// DataControl�N���X�̊�{�N���X
        /// </summary>
        public class BaseControl
        {
            private DBConnect DBConnect;
            protected SqlConnection dbControlCn;

            // BaseControl�̃R���X�g���N�^�BDBConnect�N���X�̃C���X�^���X���쐬���܂��B
            public BaseControl(string dbName)
            {
                // �f�[�^�x�[�X���I�[�v������
                DBConnect = new DBConnect(dbName);
            }

            // �f�[�^�x�[�X�ɐڑ����R�l�N�V��������Ԃ�
            public SqlConnection GetConnection()
            {
                dbControlCn = DBConnect.Cn;
                return DBConnect.Cn;
            }
        }

        public class DataControl : BaseControl
        {
            // �f�[�^�R���g���[���N���X�̃R���X�g���N�^
            public DataControl(string dbName):base(dbName)
            {
            }

            /// <summary>
            /// �f�[�^�x�[�X�ڑ�����
            /// </summary>
            public void Close()
            {
                if (dbControlCn.State == ConnectionState.Open)
                {
                    dbControlCn.Close();
                }
            }

            /// <summary>
            /// �C�ӂ�SQL�����s����
            /// </summary>
            /// <param name="tempSql">SQL��</param>
            /// <returns>���� : true, ���s : false</returns>
            public bool FreeSql(string tempSql)
            {
                bool rValue = false;

                try
                {
                    SqlCommand sCom = new SqlCommand();
                    sCom.CommandText = tempSql;
                    sCom.Connection = GetConnection();

                    //SQL�̎��s
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
            /// �f�[�^���[�_�[���擾����
            /// </summary>
            /// <param name="tempSQL">SQL��</param>
            /// <returns>�f�[�^���[�_�[</returns>
            public SqlDataReader FreeReader(string tempSQL)
            {
                SqlCommand sCom = new SqlCommand();
                sCom.CommandText = tempSQL;
                sCom.Connection = GetConnection();
                SqlDataReader dR = sCom.ExecuteReader();

                return dR;
            }

            /// <summary>
            /// �f�[�^���[�_�[���擾����
            /// </summary>
            /// <param name="tempSQL">SQL��</param>
            /// <returns>�f�[�^���[�_�[</returns>
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
            ///     �Ј��ԍ����w�肵�ĎЈ������擾���܂� </summary>
            /// <param name="sYY">
            ///     ��N</param>
            /// <param name="sMM">
            ///     ���</param>
            /// <returns>
            ///     �f�[�^���[�_�[</returns>
            /// -----------------------------------------------------------
            public SqlDataReader GetEmployeeBase(string sNo)
            {
                // SQLServer�ڑ�
                SqlDataReader dRs;
                StringBuilder sb = new StringBuilder();

                sb.Append("select CODE,NAME from SHAIN1 ");
                sb.Append("where DEL = 1 and TAIKYU = 0 and CODE = '" + sNo + "'");
                dRs = FreeReader(sb.ToString());
                return dRs;
            }

            /// -----------------------------------------------------------
            /// <summary>
            ///     �Ј������擾���܂� </summary>
            /// <param name="sYY">
            ///     ��N</param>
            /// <param name="sMM">
            ///     ���</param>
            /// <returns>
            ///     �f�[�^���[�_�[</returns>
            /// -----------------------------------------------------------
            public SqlDataReader GetEmployeeBase()
            {
                // SQLServer�ڑ�
                SqlDataReader dRs;
                StringBuilder sb = new StringBuilder();

                sb.Append("select DEL,INCODE,CODE,NAME,TAIKYU from SHAIN1 ");
                sb.Append("where DEL = 1 order by INCODE");
                dRs = FreeReader(sb.ToString());
                return dRs;
            }
        }

        /// <summary>
        /// SQLServer�f�[�^�x�[�X�ڑ��N���X
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
                    // �f�[�^�x�[�X�ڑ�������
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
