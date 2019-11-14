using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DSUN_OCR.Common;

namespace DSUN_OCR.Config
{
    public class getConfig
    {
        DSUN_IMGDataSet db = new DSUN_IMGDataSet();
        DSUN_IMGDataSetTableAdapters.環境設定TableAdapter adp = new DSUN_IMGDataSetTableAdapters.環境設定TableAdapter();

        public getConfig()
        {
            try
            {
                adp.Fill(db.環境設定);
                DSUN_IMGDataSet.環境設定Row r = db.環境設定.Single(a => a.ID == global.configKEY);

                global.cnfTifPath = r.処理済みフォルダパス;
                global.cnfOkPath = r.受け渡しデータ作成パス;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定情報取得", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
            }
        }
    }
}
