using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SZOK_OCR.Common;

namespace SZOK_OCR.OCR
{
    public partial class frmOCR : Form
    {
        public frmOCR()
        {
            InitializeComponent();

            // スキャナ画像の出力先フォルダが登録されていないとき登録します
            if (!System.IO.Directory.Exists(Properties.Settings.Default.Scan_WinoutPath))
            {
                System.IO.Directory.CreateDirectory(Properties.Settings.Default.Scan_WinoutPath);
            }

            // 画像＆CSVデータの出力先フォルダが登録されていないとき登録します
            if (!System.IO.Directory.Exists(Properties.Settings.Default.dataPath))
            {
                System.IO.Directory.CreateDirectory(Properties.Settings.Default.dataPath);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();

            // スキャナから勤務管理表を読み取りＯＣＲを実行します
            DoOCR(global.OCR_SCAN);

            this.Close();
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     WinReaderFormOCRによるＯＣＲ処理を実行します </summary>
        /// <param name="ocrMode">
        ///     入力ファイル 1:スキャン, 2:画像</param>
        /// -------------------------------------------------------------------------
        private void DoOCR(string ocrMode)
        {
            // WinReaderFormOCRを起動
            WinReaderOCR(ocrMode);

            // ファイル名のタイムスタンプを設定
            fnm = string.Format("{0:0000}", DateTime.Today.Year) +
                  string.Format("{0:00}", DateTime.Today.Month) +
                  string.Format("{0:00}", DateTime.Today.Day) +
                  string.Format("{0:00}", DateTime.Now.Hour) +
                  string.Format("{0:00}", DateTime.Now.Minute) +
                  string.Format("{0:00}", DateTime.Now.Second);

            // 連番を初期化
            dNo = 0;

            // ファイル分割処理
            LoadCsvDivide(ocrMode);
        }

        // OCRファイル名連番 : WinReader
        int dNo = 0;

        // OCRファイル名（タイムスタンプ）: WinReader
        string fnm = string.Empty;
              
        /// ----------------------------------------------------------------
        /// <summary>
        ///     WinReaderを起動してOCR処理を実施する</summary>
        /// <param name="OCRMODE">
        ///     スキャンモード 1:スキャナ, 2:イメージファイル</param>
        /// ----------------------------------------------------------------
        private void WinReaderOCR(string OCRMODE)
        {
            string JobName = string.Empty;

            // OCR認識モード毎のWinReaderJOB起動文字列
            if (OCRMODE == global.OCR_SCAN) // スキャナ
                JobName = @"""" + Properties.Settings.Default.wrHands_Job_Scan + @"""" + " /H2";
            else if (OCRMODE == global.OCR_IMAGE) // 画像
                JobName = @"""" + Properties.Settings.Default.wrHands_Job_Image + @"""" + " /H2";

            // WinReader実行ファイル
            string winReader_exe = Properties.Settings.Default.wrHands_Path + @"\" +
                Properties.Settings.Default.wrHands_Prg;

            // ProcessStartInfo の新しいインスタンスを生成する
            System.Diagnostics.ProcessStartInfo p = new System.Diagnostics.ProcessStartInfo();

            // 起動するアプリケーションを設定する
            p.FileName = winReader_exe;

            // コマンドライン引数を設定する（WinReaderのJOB起動パラメーター）
            p.Arguments = JobName;

            // WinReaderを起動します
            System.Diagnostics.Process hProcess = System.Diagnostics.Process.Start(p);

            // WinReaderが終了するまで待機する
            hProcess.WaitForExit();
        }

        /// -----------------------------------------------------------------
        /// <summary>
        ///     勤務管理表ＣＳＶデータを一枚ごとに分割する </summary>
        /// <param name="OCRMODE">
        ///     スキャンモード 1:スキャナ, 2:イメージファイル</param>
        /// -----------------------------------------------------------------
        private void LoadCsvDivide(string OCRMODE)
        {
            string imgName = string.Empty;      // 画像ファイル名
            string firstFlg = global.FLGON;
            global.pblDenNum = 0;               // 枚数を0にセット
            string[] stArrayData;               // CSVファイルを１行単位で格納する配列
            string newFnm = string.Empty;       // 新ファイル名
            string inPath = string.Empty;       // ＯＣＲ認識モードごとの入力パス
            string inFilePath = string.Empty;   // ＯＣＲ認識モードごとの入力ファイル名
            string dataMode = string.Empty;     // 勤務管理表種別（１：社員, ２：パート・アルバイト, ３：出向社員）

            // 入力ファイルパス
            if (OCRMODE == global.OCR_SCAN) // スキャナ
            {
                inPath = Properties.Settings.Default.Scan_WinoutPath;
                inFilePath = Properties.Settings.Default.Scan_WinoutPath + Properties.Settings.Default.winReaderOutFile;
            }
            else if (OCRMODE == global.OCR_IMAGE) // 画像
            {
                inPath = Properties.Settings.Default.Image_WinoutPath;
                inFilePath = Properties.Settings.Default.Image_WinoutPath + Properties.Settings.Default.winReaderOutFile;
            }

            // CSVデータの存在を確認します
            if (!System.IO.File.Exists(inFilePath)) return;

            // StreamReader の新しいインスタンスを生成する
            //入力ファイル
            System.IO.StreamReader inFile = new System.IO.StreamReader(inFilePath, Encoding.Default);

            // 読み込んだ結果をすべて格納するための変数を宣言する
            string stResult = string.Empty;
            string stBuffer;

            // 行番号
            int sRow = 0;

            // オーナーフォームを無効にする
            this.Enabled = false;

            // プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            // 勤務管理表枚数取得
            string[] t = System.IO.Directory.GetFiles(inPath, "*.tif");

            // 読み込みできる文字がなくなるまで繰り返す
            while (inFile.Peek() >= 0)
            {
                // ファイルを 1 行ずつ読み込む
                stBuffer = inFile.ReadLine();

                // カンマ区切りで分割して配列に格納する
                stArrayData = stBuffer.Split(',');

                //先頭に「*」か「#」があったら新たな伝票なのでCSVファイル作成
                if ((stArrayData[0] == "*"))
                {
                    //最初の伝票以外のとき
                    if (firstFlg == global.FLGOFF)
                    {
                        //ファイル書き出し
                        outFileWrite(stResult, inPath + imgName, newFnm);
                    }

                    //伝票枚数カウント
                    global.pblDenNum++;
                    firstFlg = global.FLGOFF;

                    // プログレス表示
                    frmP.Text = "ＯＣＲデータロード中";
                    frmP.progressValue = global.pblDenNum * 100 / t.Length;
                    frmP.ProgressStep();

                    // 伝票連番
                    dNo++;

                    // ファイル名
                    newFnm = fnm + dNo.ToString().PadLeft(3, '0');

                    // 画像ファイル名を取得
                    imgName = stArrayData[1];

                    // 勤務管理表種別を取得（１：社員、２：パート・アルバイト、３：出向社員）
                    dataMode = stArrayData[2];

                    //文字列バッファをクリア
                    stResult = string.Empty;

                    // 文字列再構成（画像ファイル名を変更する）
                    stBuffer = string.Empty;
                    for (int i = 0; i < stArrayData.Length; i++)
                    {
                        if (stBuffer != string.Empty) stBuffer += ",";

                        // 画像ファイル名を変更する
                        if (i == 1) stArrayData[i] = newFnm + ".tif"; // 画像ファイル名を変更

                        // フィールド結合
                        string sta = stArrayData[i].Trim();
                        stBuffer += sta;
                    }

                    sRow = 0;
                }
                else
                {
                    sRow++;
                }

                // 読み込んだものを追加で格納する
                stResult += (stBuffer + Environment.NewLine);                
            }

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            // 後処理
            if (global.pblDenNum > 0)
            {
                // ファイル書き出し
                outFileWrite(stResult, inPath + imgName, newFnm);

                // 入力ファイルを閉じる
                inFile.Close();

                // 入力ファイル削除 : "txtout.csv"
                Utility.FileDelete(inPath, Properties.Settings.Default.winReaderOutFile);

                // 画像ファイル削除 : "WRH***.tif"
                Utility.FileDelete(inPath, "WRH*.tif");

                // 終了表示
                MessageBox.Show(global.pblDenNum.ToString() + "件の自転車防犯登録カードを処理しました", "終了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     分割ファイルを書き出す </summary>
        /// <param name="tempResult">
        ///     書き出す文字列</param>
        /// <param name="tempImgName">
        ///     元画像ファイルパス</param>
        /// <param name="outFileName">
        ///     新ファイル名</param>
        /// -------------------------------------------------------------------------------
        private void outFileWrite(string tempResult, string tempImgName, string outFileName)
        {
            //出力ファイル
            System.IO.StreamWriter outFile = new System.IO.StreamWriter(Properties.Settings.Default.dataPath + outFileName + ".csv",
                                                    false, System.Text.Encoding.GetEncoding(932));
            // ファイル書き出し
            outFile.Write(tempResult);

            //ファイルクローズ
            outFile.Close();

            //画像ファイルをコピー
            System.IO.File.Copy(tempImgName, Properties.Settings.Default.dataPath + outFileName + ".tif");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string [] tiffile = System.IO.Directory.GetFiles(Properties.Settings.Default.imagePath, "*.tif");

            if (tiffile.Length == 0)
            {
                MessageBox.Show("勤務管理表の画像がありません", "画像ＯＣＲ処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            else
            {
                this.Hide();

                // 勤務管理表画像のＯＣＲ認識を実行します
                DoOCR(global.OCR_IMAGE);
                this.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmOCR_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }
    }
}
