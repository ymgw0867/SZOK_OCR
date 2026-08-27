using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SZOK_OCR.Common
{
    /// <summary>
    /// 防犯登録カードクラス
    /// </summary>
    public class TblRegistrationCard
    {
        public int ID { get; set; }
        public int DataCategory { get; set; }
        public string ImageFileName { get; set; }
        public string AddYear { get; set; }
        public string AddMonth { get; set; }
        public string AddDay { get; set; }
        public string Number { get; set; }
        public string VehicleIdentificationNumber { get; set; }
        public string Maker { get; set; }
        public string Color { get; set; }
        public int CarModel { get; set; }
        public string ZipCode1 { get; set; }
        public string ZipCode2 { get; set; }
        public string VehicleNumber1 { get; set; }
        public string VehicleNumber2 { get; set; }
        public string CarName { get; set; }
        public string AddressKanji { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Name { get; set; }
        public string Mobile1 { get; set; }
        public string Mobile2 { get; set; }
        public string Mobile3 { get; set; }
        public string PC { get; set; }
        public string CsvCreationDate { get; set; }
        public string Memo { get; set; }
        public DateTime UpDate { get; set; }
        public int exception { get; set; } = 0;
    }
}
