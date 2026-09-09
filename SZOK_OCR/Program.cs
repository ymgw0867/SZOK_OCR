using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using SZOK_OCR.Common;

namespace SZOK_OCR
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Mutex の新しいインスタンスを生成する (Mutex の名前にアセンブリ名を付ける)
            System.Threading.Mutex hMutex = new System.Threading.Mutex(false, Application.ProductName);

            // Mutex のシグナルを受信できるかどうか判断する
            if (hMutex.WaitOne(0, false))
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // 登録日から10年経過した防犯登録データを10年超テーブルに移動する：2026/09/09
                var master = new ClsMaster(Properties.Settings.Default.sServerName,
                                           Properties.Settings.Default.sLogin,
                                           Properties.Settings.Default.sPass,
                                           Properties.Settings.Default.sDatabase);
                DateTime date = DateTime.Today.AddYears(-25);
                master.Insert10YearsOver(date.Year * 10000 + date.Month * 100 + date.Day);

                // メインメニューを表示する
                Application.Run(new frmMainMenu());
            }
            else
            {
                // グローバル・ミューテックスによる多重起動禁止
                MessageBox.Show("このアプリケーションはすでに起動しています。2つ同時には起動できません。", "多重起動禁止");
                return;
            }

            // GC.KeepAlive メソッドが呼び出されるまで、ガベージ コレクション対象から除外される
            GC.KeepAlive(hMutex);

            // Mutex を閉じる (正しくは オブジェクトの破棄を保証する を参照)
            hMutex.Close();
        }
    }
}
