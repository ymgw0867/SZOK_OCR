using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SZOK_OCR.Common
{
    public class TblChangeNotification
    {
        public int ID { get; set; }
        public DateTime UpdateDay { get; set; }
        public string Number { get; set; }
        public string OldZipCode1 { get; set; }
        public string OldZipCode2 { get; set; }
        public string NewZipCode1 { get; set; }
        public string NewZipCode2 { get; set; }
        public string OldAddressKanji { get; set; }
        public string OldAddress1 { get; set; }
        public string OldAddress2 { get; set; }
        public string NewAddressKanji { get; set; }
        public string NewAddress1 { get; set; }
        public string NewAddress2 { get; set; }
        public string OldName { get; set; }
        public string NewName { get; set; }
        public string OldMobile1 { get; set; }
        public string OldMobile2 { get; set; }
        public string OldMobile3 { get; set; }
        public string NewMobile1 { get; set; }
        public string NewMobile2 { get; set; }
        public string NewMobile3 { get; set; }
        public string Memo { get; set; }
        public DateTime UpDate { get; set; }
    }
}
