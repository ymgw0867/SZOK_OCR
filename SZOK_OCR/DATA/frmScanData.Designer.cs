namespace SZOK_OCR.DATA
{
    partial class frmScanData
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmScanData));
            this.btnMinus = new System.Windows.Forms.Button();
            this.leadImg = new Leadtools.WinForms.RasterImageViewer();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lblNoImage = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txtTel = new System.Windows.Forms.TextBox();
            this.txtFuri = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.txtAddFuri = new System.Windows.Forms.TextBox();
            this.txtAdd = new System.Windows.Forms.TextBox();
            this.txtZip2 = new System.Windows.Forms.TextBox();
            this.txtZip1 = new System.Windows.Forms.TextBox();
            this.txtMaker = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.txtTel3 = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.txtTel2 = new System.Windows.Forms.TextBox();
            this.txtDay = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtMonth = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.txtShataiNum = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.txtTourokuNum = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.txtSharyoNum = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtCarName = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtSharyoNum2 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtStyle = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtStyleName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.txtMemo = new System.Windows.Forms.TextBox();
            this.txtYear = new System.Windows.Forms.TextBox();
            this.label55 = new System.Windows.Forms.Label();
            this.txtColor = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.label5 = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.btnPlus = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnMinus
            // 
            this.btnMinus.Image = ((System.Drawing.Image)(resources.GetObject("btnMinus.Image")));
            this.btnMinus.Location = new System.Drawing.Point(504, 634);
            this.btnMinus.Name = "btnMinus";
            this.btnMinus.Size = new System.Drawing.Size(27, 28);
            this.btnMinus.TabIndex = 116;
            this.btnMinus.TabStop = false;
            this.toolTip1.SetToolTip(this.btnMinus, "画像を縮小します");
            this.btnMinus.UseVisualStyleBackColor = true;
            this.btnMinus.Click += new System.EventHandler(this.btnMinus_Click);
            // 
            // leadImg
            // 
            this.leadImg.Location = new System.Drawing.Point(15, 16);
            this.leadImg.Name = "leadImg";
            this.leadImg.Size = new System.Drawing.Size(516, 612);
            this.leadImg.TabIndex = 123;
            this.leadImg.MouseLeave += new System.EventHandler(this.leadImg_MouseLeave);
            this.leadImg.MouseMove += new System.Windows.Forms.MouseEventHandler(this.leadImg_MouseMove);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pictureBox1.Location = new System.Drawing.Point(14, 16);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(517, 612);
            this.pictureBox1.TabIndex = 124;
            this.pictureBox1.TabStop = false;
            // 
            // lblNoImage
            // 
            this.lblNoImage.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNoImage.ForeColor = System.Drawing.SystemColors.AppWorkspace;
            this.lblNoImage.Location = new System.Drawing.Point(124, 288);
            this.lblNoImage.Name = "lblNoImage";
            this.lblNoImage.Size = new System.Drawing.Size(322, 42);
            this.lblNoImage.TabIndex = 125;
            this.lblNoImage.Text = "画像はありません";
            this.lblNoImage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Meiryo UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label8.Location = new System.Drawing.Point(99, 284);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(29, 24);
            this.label8.TabIndex = 8;
            this.label8.Text = "〒";
            // 
            // txtTel
            // 
            this.txtTel.Font = new System.Drawing.Font("Meiryo UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTel.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtTel.Location = new System.Drawing.Point(104, 494);
            this.txtTel.MaxLength = 13;
            this.txtTel.Name = "txtTel";
            this.txtTel.Size = new System.Drawing.Size(59, 32);
            this.txtTel.TabIndex = 18;
            this.txtTel.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtTel.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // txtFuri
            // 
            this.txtFuri.Font = new System.Drawing.Font("Meiryo UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtFuri.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf;
            this.txtFuri.Location = new System.Drawing.Point(104, 458);
            this.txtFuri.Name = "txtFuri";
            this.txtFuri.Size = new System.Drawing.Size(337, 32);
            this.txtFuri.TabIndex = 17;
            this.txtFuri.Leave += new System.EventHandler(this.txtMaker_Leave);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(336, 287);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(104, 26);
            this.button1.TabIndex = 14;
            this.button1.Text = "〒⇔住所変換";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtAddFuri
            // 
            this.txtAddFuri.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtAddFuri.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf;
            this.txtAddFuri.Location = new System.Drawing.Point(104, 317);
            this.txtAddFuri.Multiline = true;
            this.txtAddFuri.Name = "txtAddFuri";
            this.txtAddFuri.Size = new System.Drawing.Size(337, 64);
            this.txtAddFuri.TabIndex = 15;
            this.txtAddFuri.Leave += new System.EventHandler(this.txtMaker_Leave);
            // 
            // txtAdd
            // 
            this.txtAdd.Font = new System.Drawing.Font("Meiryo UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtAdd.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtAdd.Location = new System.Drawing.Point(104, 385);
            this.txtAdd.Multiline = true;
            this.txtAdd.Name = "txtAdd";
            this.txtAdd.Size = new System.Drawing.Size(337, 69);
            this.txtAdd.TabIndex = 16;
            // 
            // txtZip2
            // 
            this.txtZip2.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtZip2.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtZip2.Location = new System.Drawing.Point(193, 281);
            this.txtZip2.MaxLength = 4;
            this.txtZip2.Name = "txtZip2";
            this.txtZip2.Size = new System.Drawing.Size(65, 31);
            this.txtZip2.TabIndex = 13;
            this.txtZip2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtZip2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // txtZip1
            // 
            this.txtZip1.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtZip1.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtZip1.Location = new System.Drawing.Point(126, 281);
            this.txtZip1.MaxLength = 3;
            this.txtZip1.Name = "txtZip1";
            this.txtZip1.Size = new System.Drawing.Size(46, 31);
            this.txtZip1.TabIndex = 12;
            this.txtZip1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtZip1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // txtMaker
            // 
            this.txtMaker.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMaker.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf;
            this.txtMaker.Location = new System.Drawing.Point(104, 179);
            this.txtMaker.MaxLength = 10;
            this.txtMaker.Name = "txtMaker";
            this.txtMaker.Size = new System.Drawing.Size(135, 31);
            this.txtMaker.TabIndex = 6;
            this.txtMaker.Leave += new System.EventHandler(this.txtMaker_Leave);
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label18.Location = new System.Drawing.Point(283, 153);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(22, 18);
            this.label18.TabIndex = 27;
            this.label18.Text = "日";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label17.Location = new System.Drawing.Point(226, 153);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(22, 18);
            this.label17.TabIndex = 26;
            this.label17.Text = "月";
            // 
            // txtTel3
            // 
            this.txtTel3.Font = new System.Drawing.Font("Meiryo UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTel3.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtTel3.Location = new System.Drawing.Point(266, 494);
            this.txtTel3.MaxLength = 13;
            this.txtTel3.Name = "txtTel3";
            this.txtTel3.Size = new System.Drawing.Size(59, 32);
            this.txtTel3.TabIndex = 20;
            this.txtTel3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtTel3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label16.Location = new System.Drawing.Point(170, 153);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(22, 18);
            this.label16.TabIndex = 25;
            this.label16.Text = "年";
            // 
            // txtTel2
            // 
            this.txtTel2.Font = new System.Drawing.Font("Meiryo UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTel2.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtTel2.Location = new System.Drawing.Point(184, 494);
            this.txtTel2.MaxLength = 13;
            this.txtTel2.Name = "txtTel2";
            this.txtTel2.Size = new System.Drawing.Size(59, 32);
            this.txtTel2.TabIndex = 19;
            this.txtTel2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtTel2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // txtDay
            // 
            this.txtDay.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtDay.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtDay.Location = new System.Drawing.Point(249, 144);
            this.txtDay.MaxLength = 2;
            this.txtDay.Name = "txtDay";
            this.txtDay.Size = new System.Drawing.Size(32, 31);
            this.txtDay.TabIndex = 5;
            this.txtDay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtDay.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.Location = new System.Drawing.Point(165, 499);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(19, 18);
            this.label4.TabIndex = 131;
            this.label4.Text = "ー";
            // 
            // txtMonth
            // 
            this.txtMonth.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMonth.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtMonth.Location = new System.Drawing.Point(193, 144);
            this.txtMonth.MaxLength = 2;
            this.txtMonth.Name = "txtMonth";
            this.txtMonth.Size = new System.Drawing.Size(32, 31);
            this.txtMonth.TabIndex = 4;
            this.txtMonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtMonth.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label12.Location = new System.Drawing.Point(247, 499);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(19, 18);
            this.label12.TabIndex = 132;
            this.label12.Text = "ー";
            // 
            // txtShataiNum
            // 
            this.txtShataiNum.BackColor = System.Drawing.Color.White;
            this.txtShataiNum.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtShataiNum.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtShataiNum.Location = new System.Drawing.Point(104, 109);
            this.txtShataiNum.Name = "txtShataiNum";
            this.txtShataiNum.Size = new System.Drawing.Size(337, 31);
            this.txtShataiNum.TabIndex = 2;
            this.txtShataiNum.Leave += new System.EventHandler(this.txtMaker_Leave);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label14.Location = new System.Drawing.Point(174, 286);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(19, 18);
            this.label14.TabIndex = 133;
            this.label14.Text = "ー";
            // 
            // txtTourokuNum
            // 
            this.txtTourokuNum.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTourokuNum.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtTourokuNum.Location = new System.Drawing.Point(104, 76);
            this.txtTourokuNum.Name = "txtTourokuNum";
            this.txtTourokuNum.Size = new System.Drawing.Size(337, 31);
            this.txtTourokuNum.TabIndex = 1;
            this.txtTourokuNum.Leave += new System.EventHandler(this.txtMaker_Leave);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label13.Location = new System.Drawing.Point(7, 503);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(90, 18);
            this.label13.TabIndex = 13;
            this.label13.Text = "TEL／携帯：";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label15.Location = new System.Drawing.Point(8, 537);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(50, 18);
            this.label15.TabIndex = 135;
            this.label15.Text = "備考：";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label11.Location = new System.Drawing.Point(8, 467);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(61, 18);
            this.label11.TabIndex = 11;
            this.label11.Text = "フリガナ：";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label10.Location = new System.Drawing.Point(8, 409);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(50, 18);
            this.label10.TabIndex = 10;
            this.label10.Text = "住所：";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label19.Location = new System.Drawing.Point(7, 254);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(78, 18);
            this.label19.TabIndex = 138;
            this.label19.Text = "車両番号：";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label9.Location = new System.Drawing.Point(8, 338);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(89, 18);
            this.label9.TabIndex = 9;
            this.label9.Text = "住所フリガナ：";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label20.Location = new System.Drawing.Point(260, 250);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(50, 18);
            this.label20.TabIndex = 139;
            this.label20.Text = "車名：";
            // 
            // txtSharyoNum
            // 
            this.txtSharyoNum.BackColor = System.Drawing.Color.White;
            this.txtSharyoNum.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtSharyoNum.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtSharyoNum.Location = new System.Drawing.Point(104, 245);
            this.txtSharyoNum.MaxLength = 2;
            this.txtSharyoNum.Name = "txtSharyoNum";
            this.txtSharyoNum.Size = new System.Drawing.Size(33, 31);
            this.txtSharyoNum.TabIndex = 9;
            this.txtSharyoNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtSharyoNum.Leave += new System.EventHandler(this.txtMaker_Leave);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label7.Location = new System.Drawing.Point(8, 221);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(50, 18);
            this.label7.TabIndex = 7;
            this.label7.Text = "車種：";
            // 
            // txtCarName
            // 
            this.txtCarName.BackColor = System.Drawing.Color.White;
            this.txtCarName.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtCarName.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf;
            this.txtCarName.Location = new System.Drawing.Point(305, 245);
            this.txtCarName.MaxLength = 10;
            this.txtCarName.Name = "txtCarName";
            this.txtCarName.Size = new System.Drawing.Size(135, 31);
            this.txtCarName.TabIndex = 11;
            this.txtCarName.Leave += new System.EventHandler(this.txtMaker_Leave);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.Location = new System.Drawing.Point(257, 184);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 18);
            this.label6.TabIndex = 6;
            this.label6.Text = "カラー：";
            // 
            // txtSharyoNum2
            // 
            this.txtSharyoNum2.BackColor = System.Drawing.Color.White;
            this.txtSharyoNum2.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtSharyoNum2.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtSharyoNum2.Location = new System.Drawing.Point(140, 245);
            this.txtSharyoNum2.MaxLength = 5;
            this.txtSharyoNum2.Name = "txtSharyoNum2";
            this.txtSharyoNum2.Size = new System.Drawing.Size(96, 31);
            this.txtSharyoNum2.TabIndex = 10;
            this.txtSharyoNum2.Leave += new System.EventHandler(this.txtMaker_Leave);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.Location = new System.Drawing.Point(7, 153);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 18);
            this.label3.TabIndex = 3;
            this.label3.Text = "登録年月日：";
            // 
            // txtStyle
            // 
            this.txtStyle.BackColor = System.Drawing.Color.White;
            this.txtStyle.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtStyle.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtStyle.Location = new System.Drawing.Point(104, 212);
            this.txtStyle.MaxLength = 2;
            this.txtStyle.Name = "txtStyle";
            this.txtStyle.Size = new System.Drawing.Size(33, 31);
            this.txtStyle.TabIndex = 8;
            this.txtStyle.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtStyle.Leave += new System.EventHandler(this.txtStyle_Leave);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(7, 114);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(78, 18);
            this.label2.TabIndex = 2;
            this.label2.Text = "車体番号：";
            // 
            // txtStyleName
            // 
            this.txtStyleName.BackColor = System.Drawing.Color.White;
            this.txtStyleName.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtStyleName.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtStyleName.Location = new System.Drawing.Point(140, 212);
            this.txtStyleName.MaxLength = 5;
            this.txtStyleName.Name = "txtStyleName";
            this.txtStyleName.ReadOnly = true;
            this.txtStyleName.Size = new System.Drawing.Size(300, 31);
            this.txtStyleName.TabIndex = 144;
            this.txtStyleName.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(7, 81);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 18);
            this.label1.TabIndex = 1;
            this.label1.Text = "登録番号：";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label21.Location = new System.Drawing.Point(111, 153);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(26, 18);
            this.label21.TabIndex = 145;
            this.label21.Text = "20";
            // 
            // txtMemo
            // 
            this.txtMemo.Font = new System.Drawing.Font("Meiryo UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMemo.Location = new System.Drawing.Point(104, 531);
            this.txtMemo.Multiline = true;
            this.txtMemo.Name = "txtMemo";
            this.txtMemo.Size = new System.Drawing.Size(336, 30);
            this.txtMemo.TabIndex = 21;
            // 
            // txtYear
            // 
            this.txtYear.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtYear.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtYear.Location = new System.Drawing.Point(136, 144);
            this.txtYear.MaxLength = 4;
            this.txtYear.Name = "txtYear";
            this.txtYear.Size = new System.Drawing.Size(32, 31);
            this.txtYear.TabIndex = 3;
            this.txtYear.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtYear.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            // 
            // label55
            // 
            this.label55.AutoSize = true;
            this.label55.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label55.Location = new System.Drawing.Point(8, 184);
            this.label55.Name = "label55";
            this.label55.Size = new System.Drawing.Size(64, 18);
            this.label55.TabIndex = 146;
            this.label55.Text = "メーカー：";
            // 
            // txtColor
            // 
            this.txtColor.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtColor.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf;
            this.txtColor.Location = new System.Drawing.Point(305, 179);
            this.txtColor.Name = "txtColor";
            this.txtColor.Size = new System.Drawing.Size(135, 31);
            this.txtColor.TabIndex = 7;
            this.txtColor.Leave += new System.EventHandler(this.txtMaker_Leave);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.radioButton2);
            this.panel1.Controls.Add(this.radioButton1);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.txtColor);
            this.panel1.Controls.Add(this.label55);
            this.panel1.Controls.Add(this.txtYear);
            this.panel1.Controls.Add(this.txtMemo);
            this.panel1.Controls.Add(this.label21);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.txtStyleName);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.txtStyle);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.txtSharyoNum2);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.txtCarName);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.txtSharyoNum);
            this.panel1.Controls.Add(this.label20);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.label19);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.label15);
            this.panel1.Controls.Add(this.label13);
            this.panel1.Controls.Add(this.txtTourokuNum);
            this.panel1.Controls.Add(this.label14);
            this.panel1.Controls.Add(this.txtShataiNum);
            this.panel1.Controls.Add(this.label12);
            this.panel1.Controls.Add(this.txtMonth);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.txtDay);
            this.panel1.Controls.Add(this.txtTel2);
            this.panel1.Controls.Add(this.label16);
            this.panel1.Controls.Add(this.txtTel3);
            this.panel1.Controls.Add(this.label17);
            this.panel1.Controls.Add(this.label18);
            this.panel1.Controls.Add(this.txtMaker);
            this.panel1.Controls.Add(this.txtZip1);
            this.panel1.Controls.Add(this.txtZip2);
            this.panel1.Controls.Add(this.txtAdd);
            this.panel1.Controls.Add(this.txtAddFuri);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.txtFuri);
            this.panel1.Controls.Add(this.txtTel);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Location = new System.Drawing.Point(537, 16);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(459, 612);
            this.panel1.TabIndex = 0;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(174, 51);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(52, 21);
            this.radioButton2.TabIndex = 154;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "原付";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(104, 51);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(65, 21);
            this.radioButton1.TabIndex = 153;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "自転車";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.Location = new System.Drawing.Point(8, 53);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(81, 18);
            this.label5.TabIndex = 152;
            this.label5.Text = "カード種類：";
            // 
            // linkLabel1
            // 
            this.linkLabel1.Font = new System.Drawing.Font("Meiryo UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.linkLabel1.Image = ((System.Drawing.Image)(resources.GetObject("linkLabel1.Image")));
            this.linkLabel1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.linkLabel1.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkLabel1.Location = new System.Drawing.Point(938, 641);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(58, 21);
            this.linkLabel1.TabIndex = 127;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "終了";
            this.linkLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // btnPlus
            // 
            this.btnPlus.Image = ((System.Drawing.Image)(resources.GetObject("btnPlus.Image")));
            this.btnPlus.Location = new System.Drawing.Point(478, 634);
            this.btnPlus.Name = "btnPlus";
            this.btnPlus.Size = new System.Drawing.Size(27, 28);
            this.btnPlus.TabIndex = 131;
            this.btnPlus.TabStop = false;
            this.toolTip1.SetToolTip(this.btnPlus, "画像を拡大します");
            this.btnPlus.UseVisualStyleBackColor = true;
            this.btnPlus.Click += new System.EventHandler(this.btnPlus_Click_1);
            // 
            // frmScanData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1010, 683);
            this.Controls.Add(this.btnPlus);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.lblNoImage);
            this.Controls.Add(this.leadImg);
            this.Controls.Add(this.btnMinus);
            this.Controls.Add(this.pictureBox1);
            this.Font = new System.Drawing.Font("Meiryo UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "frmScanData";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "防犯登録カードデータ作成";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmCorrect_FormClosing);
            this.Load += new System.EventHandler(this.frmCorrect_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btnMinus;
        private Leadtools.WinForms.RasterImageViewer leadImg;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label lblNoImage;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtTel;
        private System.Windows.Forms.TextBox txtFuri;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtAddFuri;
        private System.Windows.Forms.TextBox txtAdd;
        private System.Windows.Forms.TextBox txtZip2;
        private System.Windows.Forms.TextBox txtZip1;
        private System.Windows.Forms.TextBox txtMaker;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TextBox txtTel3;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox txtTel2;
        private System.Windows.Forms.TextBox txtDay;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtMonth;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtShataiNum;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox txtTourokuNum;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.TextBox txtSharyoNum;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtCarName;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtSharyoNum2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtStyle;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtStyleName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.TextBox txtMemo;
        private System.Windows.Forms.TextBox txtYear;
        private System.Windows.Forms.Label label55;
        private System.Windows.Forms.TextBox txtColor;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Button btnPlus;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.Label label5;
    }
}

