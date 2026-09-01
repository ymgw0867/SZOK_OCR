using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SZOK_OCR.Common
{
    public class TblCSVCreationHistory
    {
        public int ID { get; set; }
        public DateTime CreationDate { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime FinishDate { get; set; }
        public int Outputs { get; set; }
        public string PC { get; set; }
        public string Memo { get; set; }
    }
}
