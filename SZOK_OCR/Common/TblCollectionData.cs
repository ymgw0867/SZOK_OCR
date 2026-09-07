using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SZOK_OCR.Common
{
    /// <summary>
    /// 回収データ
    /// </summary>
    public class TblCollectionData
    {
        /// <summary>
        /// ID
        /// </summary>
        public int ID { get; set; }
        /// <summary>
        /// 出庫ID
        /// </summary>
        public int ShippingID { get; set; }
        /// <summary>
        /// 回収日
        /// </summary>
        public DateTime CollectionDate { get; set; }
        /// <summary>
        /// 登録番号
        /// </summary>
        public int Number { get; set; }
        /// <summary>
        /// 防犯登録ID
        /// </summary>
        public int CardID { get; set; }
        /// <summary>
        /// SCANID
        /// </summary>
        public int ScanID { get; set; }
        /// <summary>
        /// 更新日時
        /// </summary>
        public DateTime Update { get; set; }

    }
}
