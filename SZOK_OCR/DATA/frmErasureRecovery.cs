using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SZOK_OCR.Common;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace SZOK_OCR.DATA
{
    public partial class frmErasureRecovery : Form
    {
        public frmErasureRecovery(int id, string cpa, string eraDate, string name, string address)
        {
            InitializeComponent();
            _id = id;
            _cpa = cpa;
            _eraDate = eraDate;
            _name = name;
            _address = address;
            status = false;
        }

        public bool status { get; set; }

        int _id;
        string _cpa;
        string _eraDate;
        string _name;
        string _address;

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmErasure_Load(object sender, EventArgs e)
        {
            ShowData(_id);
        }

        /// <summary>
        /// 指定されたIDのデータを表示する
        /// </summary>
        /// <param name="iX">表示するデータのID</param>
        private void ShowData(int iX)
        {
            // 情報表示
            lblNumber.Text = _cpa;
            lblEraDate.Text = _eraDate;
            lblName.Text = _name;
            lblAddress.Text = _address;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 確認メッセージ
            if (MessageBox.Show("抹消を取り消して防犯登録データに復元します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // 抹消取消（復元）処理
            var master = new ClsMaster(Properties.Settings.Default.sServerName, Properties.Settings.Default.sLogin,
                                   Properties.Settings.Default.sPass, Properties.Settings.Default.sDatabase);
            master.ErasureRecovery(_id);

            // ログ記録
            RecoveryLogCreate(_cpa, _name);

            // 終了メッセージ
            MessageBox.Show("抹消取消（復元）処理が完了しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);

            status = true;
            Close();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 復元操作の記録を CSV 形式のログファイルに作成および追記する。
        /// </summary>
        /// <remarks>ログファイルが存在しない場合は Properties.Settings.Default.sLogPath
        /// で指定されたファイルを作成し、先頭にヘッダー行（復元日,CPA,氏名）を書き込む。各ログ行は yyyy/MM/dd HH:mm:ss 形式のタイムスタンプ、CPA、氏名をカンマ区切りで追記する。</remarks>
        /// <param name="cpa">復元対象の CPA を表す文字列。</param>
        /// <param name="name">復元対象者の氏名。</param>
        private void RecoveryLogCreate(string cpa, string name)
        {
            if (!System.IO.File.Exists(Properties.Settings.Default.sLogPath))
            {
                System.IO.File.Create(Properties.Settings.Default.sLogPath);

                // CSVファイルのヘッダー行を作成する
                var header = "復元日,CPA,氏名";
                System.IO.File.WriteAllText(Properties.Settings.Default.sLogPath, header + Environment.NewLine);
            }

            // CSVファイルに復元情報を追記する
            var logEntry = $"{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")},{cpa},{name}";
            System.IO.File.AppendAllText(Properties.Settings.Default.sLogPath, logEntry + Environment.NewLine);
        }
    }
}
    