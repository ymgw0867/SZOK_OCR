using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SZOK_OCR.Common
{
    public class TblShippingout
    {
        /// <summary>
        /// ID
        /// </summary>
        public int ID { get; set; }
        /// <summary>
        /// 出荷日
        /// </summary>
        public DateTime ShippingDate { get; set; }
        /// <summary>
        /// 店舗番号
        /// </summary>
        public int ShopNumber { get; set; }
        /// <summary>
        /// 店舗名
        /// </summary>
        public string ShopName { get; set; }
        /// <summary>
        /// 部数
        /// </summary>
        public int Copies { get; set; }
        /// <summary>
        /// 開始番号
        /// </summary>
        public int StartNumber { get; set; }
        /// <summary>
        /// 終了番号
        /// </summary>
        public int FinishNumber { get; set; }
        /// <summary>
        /// 売上
        /// </summary>
        public int Sales { get; set; }

    }
}
