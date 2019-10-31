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
    public partial class frmCorrect : Form
    {
        SZOKDataSetTableAdapters.OCR防犯TableAdapter ocrAdp = new SZOKDataSetTableAdapters.OCR防犯TableAdapter();
        SZOKDataSetTableAdapters.自転車防犯登録TableAdapter dataAdp = new SZOKDataSetTableAdapters.自転車防犯登録TableAdapter();
        SZOKDataSet dts = new SZOKDataSet();

        // 検索ＩＤ
        string dID = string.Empty;

        /// <summary>
        ///     カレントデータRowsインデックス</summary>
        int cI = 0;

        public frmCorrect(string sID)
        {
            InitializeComponent();

            if (sID == string.Empty)    // ＯＣＲデータを新規登録する
            {
                // 画像パス取得
                global.pblImagePath = Properties.Settings.Default.dataPath;

                // ＯＣＲデータ読み込み   
                ocrAdp.Fill(dts.OCR防犯);
            }
            else　// 登録済みデータの検索及び編集
            {
                dID = sID;
                dataAdp.Fill(dts.自転車防犯登録);
            }
        }

        private void frmCorrect_Load(object sender, EventArgs e)
        {
            this.pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            //元号を取得
            label15.Text = Properties.Settings.Default.gengou;

            // globalクラスインスタンス
            global gl = new global();

            // メーカーをコンボボックスにロード
            cmbMaker.DataSource = gl.arrMaker;

            // 塗色をコンボボックスにロード
            cmbColor.DataSource = gl.arrColor;

            // 車種をコンボボックスにロード
            cmbStyle.DataSource = gl.arrStyle;

            // 勤務データ登録
            if (dID == string.Empty)
            {
                // CSVデータをMDBへ読み込みます
                GetCsvDataToMDB();

                // データセットへデータを読み込みます
                ocrAdp.Fill(dts.OCR防犯);
                //getDataSet();

                // データテーブル件数カウント
                if (dts.OCR防犯.Count == 0)
                {
                    MessageBox.Show("対象となるＯＣＲ自転車防犯登録データがありません", "ＯＣＲ自転車防犯登録データ登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //終了処理
                    Environment.Exit(0);
                }
            }

            // キャプション
            this.Text = "ＯＣＲ自転車防犯データ登録";

            // 編集作業、過去データ表示の判断
            if (dID == string.Empty) // パラメータのヘッダIDがないときは編集作業
            {
                // 最初のレコードを表示
                cI = 0;
                showOcrData(cI);
            }

            // tagを初期化
            this.Tag = string.Empty;
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVデータをMDBへインサートする</summary>
        ///----------------------------------------------------------------------------
        private void GetCsvDataToMDB()
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
            ocr.CsvToMdb(Properties.Settings.Default.dataPath, frmP, dts);

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     伝票画像表示 </summary>
        /// <param name="iX">
        ///     現在の伝票</param>
        /// <param name="tempImgName">
        ///     画像名</param>
        /// ------------------------------------------------------------------------------
        public void ShowImage(string tempImgName)
        {
            //修正画面へ組み入れた画像フォームの表示    
            //画像の出力が無い場合は、画像表示をしない。
            if (tempImgName == string.Empty)
            {
                leadImg.Visible = false;
                lblNoImage.Visible = false;
                global.pblImagePath = string.Empty;
                return;
            }

            //画像ファイルがあるとき表示
            if (System.IO.File.Exists(tempImgName))
            {
                lblNoImage.Visible = false;
                leadImg.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = true;
                btnMinus.Enabled = true;

                //画像ロード
                Leadtools.Codecs.RasterCodecs.Startup();
                Leadtools.Codecs.RasterCodecs cs = new Leadtools.Codecs.RasterCodecs();

                // 描画時に使用される速度、品質、およびスタイルを制御します。 
                Leadtools.RasterPaintProperties prop = new Leadtools.RasterPaintProperties();
                prop = Leadtools.RasterPaintProperties.Default;
                prop.PaintDisplayMode = Leadtools.RasterPaintDisplayModeFlags.Resample;
                leadImg.PaintProperties = prop;

                leadImg.Image = cs.Load(tempImgName, 0, Leadtools.Codecs.CodecsLoadByteOrder.BgrOrGray, 1, 1);

                //画像表示倍率設定
                if (global.miMdlZoomRate == 0f)
                {
                    leadImg.ScaleFactor *= global.ZOOM_RATE;
                }
                else
                {
                    leadImg.ScaleFactor *= global.miMdlZoomRate;
                }

                //画像のマウスによる移動を可能とする
                leadImg.InteractiveMode = Leadtools.WinForms.RasterViewerInteractiveMode.Pan;

                // グレースケールに変換
                Leadtools.ImageProcessing.GrayscaleCommand grayScaleCommand = new Leadtools.ImageProcessing.GrayscaleCommand();
                grayScaleCommand.BitsPerPixel = 8;
                grayScaleCommand.Run(leadImg.Image);
                leadImg.Refresh();

                cs.Dispose();
                Leadtools.Codecs.RasterCodecs.Shutdown();
                //global.pblImagePath = tempImgName;
            }
            else
            {
                //画像ファイルがないとき
                lblNoImage.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = false;
                btnMinus.Enabled = false;

                leadImg.Visible = false;
                //global.pblImagePath = string.Empty;
            }
        }

    }
}
