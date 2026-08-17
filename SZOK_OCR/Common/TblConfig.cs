using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SZOK_OCR.Common
{
    public class TblConfig
    {
        public int ID { get; set; }
        public int Year { get; set; }
        public int Month { get; set; }
        public int DataSaveMonth { get; set; }
        public string ZipCodePath { get; set; }
        public DateTime UpDate { get; set; }
    }
}
