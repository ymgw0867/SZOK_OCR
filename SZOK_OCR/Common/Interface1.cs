using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace SZOK_OCR.Common
{
    public interface IMaster
    {
        T GetData<T>(string id, SqlConnection conn);

        //List<T> Read<T>(string s);

        //List<T> Read<T>();

        void UpDate<T>(T Object, SqlConnection conn);

        //int Count<T>();

        //int Count<T>(string id);

        void Insert<T>(T Object, SqlConnection conn);

        //bool Delete<T>(string sql);
        //bool Delete<T>();
    }
}
