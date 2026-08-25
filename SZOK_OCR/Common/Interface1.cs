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
        T GetData<T>(string id);

        //List<T> Read<T>(string s);

        List<T> Read<T>(ScandataParameter param);

        //List<T> Read<T>();

        void UpDate<T>(T Object);

        int Count<T>();

        int Count<T>(string id);

        void Insert<T>(List<T> Object);

        //bool Delete<T>(string sql);
        //bool Delete<T>();
    }
}
