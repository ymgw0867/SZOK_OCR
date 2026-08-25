using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SZOK_OCR.Common
{
    // スキャンデータテーブル
    public class TblScandata
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
        public string Label { get; set; }
        public string Person { get; set; }
    }

    public class ScandataParameter
    {
        public int? DataCategory { get; set; }   // データ区分
        public string AddYear { get; set; }     // 登録年
        public string AddMonth { get; set; }    // 登録月
        public string AddDay { get; set; }      // 登録日
        public string Number { get; set; }      // 登録番号
        public string VehicleIdentificationNumber { get; set; } // 車両番号
        public string Maker { get; set; }       // メーカー
        public string Color { get; set; }       // 色
        public int? CarModel { get; set; }    // 車種
        public string ZipCode1 { get; set; }    // 郵便番号1
        public string ZipCode2 { get; set; }    // 郵便番号2
        public string VehicleNumber1 { get; set; }  // 車両番号1
        public string VehicleNumber2 { get; set; }  // 車両番号2
        public string CarName { get; set; }     // 車名
        public string Address1 { get; set; }    // 住所1
        public string Address2 { get; set; }    // 住所2
        public string Name { get; set; }    // 氏名
        public string Mobile1 { get; set; } // 携帯電話1
        public string Mobile2 { get; set; } // 携帯電話2
        public string Mobile3 { get; set; } // 携帯電話3
        public string Label { get; set; }   // ラベル
        public string Person { get; set; }  // 担当者
    }
}
