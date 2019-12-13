﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Leadtools;
using Leadtools.Codecs;
using Leadtools.ImageProcessing;
//using Leadtools.ImageProcessing.Core;
using System.Data.OleDb;
using SZOK_OCR.Common;
using Excel = Microsoft.Office.Interop.Excel;

namespace SZOK_OCR.OCR
{
    public partial class frmOCR : Form
    {
        public frmOCR()
        {
            InitializeComponent();
        }
        
        string _outPC = string.Empty;

        private void button1_Click(object sender, EventArgs e)
        {
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     帳票認識ライブラリ V8.0.3 による認識処理実行
        /// </summary>
        ///----------------------------------------------------------------
        private void wrhs803LibOCR(string jobName)
        {
            // ファイル名のタイムスタンプを設定
            string fnm = string.Format("{0:0000}", DateTime.Today.Year) + 
                         string.Format("{0:00}", DateTime.Today.Month) +
                         string.Format("{0:00}", DateTime.Today.Day) +
                         string.Format("{0:00}", DateTime.Now.Hour) +
                         string.Format("{0:00}", DateTime.Now.Minute) +
                         string.Format("{0:00}", DateTime.Now.Second);

            int sNum = 0;
            int sOK = 0;
            int sNG = 0;
            int ret = 0;

            try
            {
                // オーナーフォームを無効にする
                this.Enabled = false;

                // プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = this;
                frmP.Show();

                // 処理する画像数を取得
                int t = System.IO.Directory.GetFiles(Properties.Settings.Default.trayPath, "*.tif").Count();

                // 順番に認識処理を実行
                foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.trayPath, "*.tif"))
                {
                    // 画像数カウント
                    sNum++;

                    // プログレス表示
                    frmP.Text = "OCR認識中です ... " + sNum.ToString() + "/" + t.ToString();
                    frmP.progressValue = sNum * 100 / t;
                    frmP.ProgressStep();

                    // 標準パターンの読み込み
                    ret = FormRecog.OcrPatternLoad(Properties.Settings.Default.ocrPatternLoadPath);
                    
                    // パターン読み込みに成功したとき
                    if (ret > 0)
                    {
                        // 帳票認識ライブラリの制御内容を設定
                        FormRecog.OcrSetStatus(5, 0);   // 強制認識制御（強制認識しない）
                        FormRecog.OcrSetStatus(28, 0);  // フリーピッチ手書認識フィールド（「数」のみ）（０：接触なし）
                        FormRecog.OcrSetStatus(29, 0);  // フリーピッチ手書認識フィールド（「数」のみ以外）（０：結合しない）

                        // 認識結果出力イメージファイル
                        StringBuilder outimage = new StringBuilder(256);
                        outimage.Append(Properties.Settings.Default.wrOutPath + System.IO.Path.GetFileName(files));

                        // 認識結果出力テキストファイル
                        StringBuilder outtext = new StringBuilder(256);
                        outtext.Append(Properties.Settings.Default.wrOutPath + System.IO.Path.GetFileNameWithoutExtension(files) + ".csv");

                        // 認識結果 構造体
                        FormRecog.FORM_RECOG_DATA dt = new FormRecog.FORM_RECOG_DATA();

                        // 認識処理を開始 : LearnFlagをtrue 2016/06/06
                        ret = FormRecog.OcrFormRecogStart(jobName, files, outimage, outtext, ref dt, false, false);

                        // 認識成功のとき
                        if (ret > 0)
                        {
                            // 認識結果のメモリ解放
                            ret = FormRecog.OcrFormStructFree(ref dt);

                            // 認識終了
                            ret = FormRecog.OcrFormRecogEnd();

                            // 出力されたイメージファイルとテキストファイルのリネーム処理を行います
                            // READフォルダ → DATAフォルダ
                            string inCsvFile = Properties.Settings.Default.wrOutPath +
                                               Properties.Settings.Default.wrReaderOutFile;
                            string newFileName = Properties.Settings.Default.dataPath + fnm + sNum.ToString().PadLeft(3, '0');
                            wrhOutFileRename(inCsvFile, newFileName);

                            // カウント
                            sOK++;
                        }
                        else
                        {
                            //MessageBox.Show("OCR認識開始に失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                            // NGフォルダがあるか？なければ作成する（TIFフォルダ）
                            if (!System.IO.Directory.Exists(Properties.Settings.Default.ngPath))
                            {
                                System.IO.Directory.CreateDirectory(Properties.Settings.Default.ngPath);
                            }

                            // 移動先NGフォルダパス
                            string toImg = Properties.Settings.Default.ngPath + System.IO.Path.GetFileName(files);

                            // 同名ファイルが既に登録済みのときは削除する
                            if (System.IO.File.Exists(toImg)) System.IO.File.Delete(toImg);

                            // NG画像をコピーする
                            if (System.IO.File.Exists(files)) System.IO.File.Copy(files, toImg);
                            
                            // NGカウント
                            sNG++;
                        }
                    }
                    else
                    {
                        MessageBox.Show("OCR標準パターンの読み込みに失敗しました", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }

                // いったんオーナーをアクティブにする
                this.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                this.Enabled = true;

                // 終了表示
                string msg = sNum.ToString() + "件のOCR認識処理を行いました" + Environment.NewLine + Environment.NewLine;
                msg += "OK件数 ： " + sOK.ToString() + Environment.NewLine;
                msg += "NG件数 ： " + sNG.ToString();
                MessageBox.Show(msg, "OCR認識結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
               
                // TRAYフォルダの全てのtifファイルを削除します
                foreach (var files in System.IO.Directory.GetFiles(Properties.Settings.Default.trayPath, "*.tif"))
                {
                    System.IO.File.Delete(files);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
        }

        // OCRファイル名連番 : WinReader
        //int dNo = 0;
        
        /// -----------------------------------------------------------------
        /// <summary>
        ///     CSVファイルと画像ファイルの名前を日付スタンプに変更する </summary>
        /// <param name="readFilePath">
        ///     入力CSVファイル名(フルパス）</param>
        /// <param name="newFnm">
        ///     新ファイル名（フルパス・但し拡張子なし）</param>
        /// -----------------------------------------------------------------
        private void wrhOutFileRename(string readFilePath, string newFnm)
        {
            string imgName = string.Empty;      // 画像ファイル名
            string[] stArrayData;               // CSVファイルを１行単位で格納する配列
            string inFilePath = string.Empty;   // ＯＣＲ認識モードごとの入力ファイル名

            // CSVデータの存在を確認します
            if (!System.IO.File.Exists(readFilePath)) return;

            // StreamReader の新しいインスタンスを生成する
            //入力ファイル
            System.IO.StreamReader inFile = new System.IO.StreamReader(readFilePath, Encoding.Default);

            // 読み込んだ結果をすべて格納するための変数を宣言する
            string stResult = string.Empty;
            string stBuffer;

            // 読み込みできる文字がなくなるまで繰り返す
            while (inFile.Peek() >= 0)
            {
                // ファイルを 1 行ずつ読み込む
                stBuffer = inFile.ReadLine();

                // カンマ区切りで分割して配列に格納する
                stArrayData = stBuffer.Split(',');

                //先頭に「*」か「#」があったらヘッダー情報
                if ((stArrayData[0] == "*"))
                {
                    //文字列バッファをクリア
                    stResult = string.Empty;

                    // 文字列再構成（画像ファイル名を変更する）
                    stBuffer = string.Empty;
                    for (int i = 0; i < stArrayData.Length; i++)
                    {
                        if (stBuffer != string.Empty)
                        {
                            stBuffer += ",";
                        }

                        // 画像ファイル名を変更する
                        if (i == 1)
                        {
                            stArrayData[i] = System.IO.Path.GetFileName(newFnm) + ".tif"; // 画像ファイル名を変更
                        }

                        // 登録番号中のスペースを除去して"CPA"を先頭に追加する 2016/06/15
                        if (i == 2)
                        {
                            stArrayData[i] = global.DATA_CPA + stArrayData[i].Replace(" ", "");
                        }

                        // フィールド結合
                        //string sta = stArrayData[i].Trim();
                        //stBuffer += sta;

                        stBuffer += stArrayData[i];
                    }
                }

                // 読み込んだものを追加で格納する
                stResult += (stBuffer + Environment.NewLine);
            }

            // CSVファイル書き出し
            System.IO.StreamWriter outFile = new System.IO.StreamWriter(newFnm + ".csv",
                                                    false, System.Text.Encoding.GetEncoding(932));
            outFile.Write(stResult);

            // 出力ファイルを閉じる
            outFile.Close();

            // 入力ファイルを閉じる
            inFile.Close();

            // 入力ファイル削除 : "txtout.csv"
            string inPath = System.IO.Path.GetDirectoryName(readFilePath);
            Utility.FileDelete(inPath, Properties.Settings.Default.wrReaderOutFile);

            // 画像ファイルをリネーム
            System.IO.File.Move(Properties.Settings.Default.wrOutPath + "WRH00001.tif", newFnm + ".tif");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmOCR_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        ///------------------------------------------------------------------------------
        /// <summary>
        ///     マルチフレームの画像ファイルを頁ごとに分割する </summary>
        /// <param name="InPath">
        ///     画像ファイル入力パス</param>
        /// <param name="outPath">
        ///     分割後出力パス</param>
        /// <returns>
        ///     true:分割を実施, false:分割ファイルなし</returns>
        ///------------------------------------------------------------------------------
        private bool MultiTif(string InPath, string outPath)
        {
            //スキャン出力画像を確認
            if (System.IO.Directory.GetFiles(InPath, "*.tif").Count() == 0)
            {
                MessageBox.Show("ＯＣＲ変換処理対象の画像ファイルが指定フォルダ " + InPath + " に存在しません", "スキャン画像確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            // 出力先フォルダがなければ作成する
            if (System.IO.Directory.Exists(outPath) == false)
            {
                System.IO.Directory.CreateDirectory(outPath);
            }

            // 出力先フォルダ内の全てのファイルを削除する（通常ファイルは存在しないが例外処理などで残ってしまった場合に備えて念のため）
            foreach (string files in System.IO.Directory.GetFiles(outPath, "*"))
            {
                System.IO.File.Delete(files);
            }

            RasterCodecs.Startup();
            RasterCodecs cs = new RasterCodecs();

            int _pageCount = 0;
            string fnm = string.Empty;

            // マルチTIFを分解して画像ファイルをTRAYフォルダへ保存する
            foreach (string files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                // 画像読み出す
                RasterImage leadImg = cs.Load(files, 0, CodecsLoadByteOrder.BgrOrGray, 1, -1);

                // 頁数を取得
                int _fd_count = leadImg.PageCount;

                // 頁ごとに読み出す
                for (int i = 1; i <= _fd_count; i++)
                {
                    // ファイル名（日付時間部分）
                    string fName = string.Format("{0:0000}", DateTime.Today.Year) +
                            string.Format("{0:00}", DateTime.Today.Month) +
                            string.Format("{0:00}", DateTime.Today.Day) +
                            string.Format("{0:00}", DateTime.Now.Hour) +
                            string.Format("{0:00}", DateTime.Now.Minute) +
                            string.Format("{0:00}", DateTime.Now.Second);

                    // ファイル名設定
                    _pageCount++;
                    fnm = outPath + fName + string.Format("{0:000}", _pageCount) + ".tif";

                    // 画像保存
                    cs.Save(leadImg, fnm, RasterImageFormat.Tif, 0, i, i, 1, CodecsSavePageMode.Insert);                    
                }
            }

            // InPathフォルダの全てのtifファイルを削除する
            foreach (var files in System.IO.Directory.GetFiles(InPath, "*.tif"))
            {
                System.IO.File.Delete(files);
            }

            RasterCodecs.Shutdown();
            return true;
        }

        private void frmOCR_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最少サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);
            
            // ラベル（日付時間部分） 2019/11/12
            label3.Text = string.Format("{0:0000}", DateTime.Now.Year) +
                    string.Format("{0:00}", DateTime.Now.Month) +
                    string.Format("{0:00}", DateTime.Now.Day) +
                    string.Format("{0:00}", DateTime.Now.Hour) +
                    string.Format("{0:00}", DateTime.Now.Minute) +
                    string.Format("{0:00}", DateTime.Now.Second);

            // 処理担当者 2019/11/12
            txtName.Text = "";
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 入力画像があるか？
            var tifCnt = System.IO.Directory.GetFiles(Properties.Settings.Default.scanPath, "*.tif").Count();

            if (tifCnt == 0)
            {
                MessageBox.Show("防犯登録カード画像がありません", "画像なし", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 2019/11/12
            if (txtName.Text.Trim() == string.Empty)
            {
                MessageBox.Show("処理担当者名が未入力です", "入力不備", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtName.Focus();
                return;
            }


            if (MessageBox.Show("OCR認識を実行します。" + Environment.NewLine + "よろしいですか", "OCR実行確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                return;

            string jobname = Properties.Settings.Default.wrHands_JobName;
            string msg = string.Empty;

            this.Hide();

            // マルチTiff画像をシングルtifに分解する(SCANフォルダ → TRAYフォルダ)
            if (!MultiTif(Properties.Settings.Default.scanPath, Properties.Settings.Default.trayPath))
            {
                this.Show();
                return;
            }

            //// ノイズ除去
            //imgMedian(Properties.Settings.Default.trayPath, 1);

            // 帳票ライブラリV8.0.3によるOCR認識実行
            wrhs803LibOCR(jobname);

            // ＯＣＲ認識結果をSCAN_DATAに書き出す : 2019/11/12
            CsvToScanData();

            // 画像ファイルをSCANDATAフォルダへ移動 : 2019/11/15
            foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.dataPath, "*.tif"))
            {
                System.IO.File.Move(files, Properties.Settings.Default.scanDataPath + System.IO.Path.GetFileName(files));
            }

            // ラベル発行：2019/11/12
            LabelReport(label3.Text, txtName.Text);

            // フォームを閉じる
            this.Close();
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     防犯登録カードラベル印刷 </summary>
        /// <param name="sLabel">
        ///     ラベル文字列</param>
        /// <param name="sName">
        ///     処理担当者名</param>
        ///---------------------------------------------------------------
        private void LabelReport(string sLabel, string sName)
        {
            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsLabelPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                try
                {
                    oxlsSheet.Cells[10, 2] = sLabel;                 
                    oxlsSheet.Cells[18, 5] = DateTime.Today.ToShortDateString();
                    oxlsSheet.Cells[19, 5] = sName;

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    oXls.Visible = true;

                    //印刷
                    oxlsSheet.PrintPreview(true);

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "ラベル印刷エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                finally
                {

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();

                    // COM オブジェクトの参照カウントを解放する 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "ラベル印刷エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }


        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVデータをSCAN_Dataへインサートする</summary>
        ///----------------------------------------------------------------------------
        private void CsvToScanData()
        {
            // CSVファイルがなければ終了
            if (System.IO.Directory.GetFiles(Properties.Settings.Default.dataPath, "*.csv").Count() == 0)
            {
                return;
            }

            // オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            // OCRのCSVデータをMDBへ取り込む
            OCRData ocr = new OCRData();
            ocr.CsvToMdb(Properties.Settings.Default.dataPath, frmP, label3.Text, txtName.Text);

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;
        }

    }
}
