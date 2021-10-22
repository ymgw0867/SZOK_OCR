using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SZOK_OCR.Common
{
    public class ClsShukko
    {
        public int ID { get; set; }
        public int UCode { get; set; }
        public string User { get; set; }
        public string ShukkoDate { get; set; }
        public int StartNumber { get; set; }
        public int EndNumber { get; set; }
        public int Suu { get; set; }
        public int Kaishu { get; set; }
        public int Zan { get; set; }
        public DateTime kaishuLimitDate { get; set; }
    }
}
