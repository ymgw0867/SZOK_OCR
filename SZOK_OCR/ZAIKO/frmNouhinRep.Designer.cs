namespace SZOK_OCR.ZAIKO
{
    partial class frmNouhinRep
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmNouhinRep));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label4 = new System.Windows.Forms.Label();
            this.txtTenName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtBusu = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtSNum = new System.Windows.Forms.TextBox();
            this.txtENum = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txtDaikin = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.txtDenNum = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(19, 46);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(887, 362);
            this.dataGridView1.TabIndex = 5;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.Location = new System.Drawing.Point(51, 52);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(49, 14);
            this.label4.TabIndex = 6;
            this.label4.Text = "店名：";
            // 
            // txtTenName
            // 
            this.txtTenName.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            this.txtTenName.Location = new System.Drawing.Point(100, 47);
            this.txtTenName.Name = "txtTenName";
            this.txtTenName.Size = new System.Drawing.Size(781, 22);
            this.txtTenName.TabIndex = 7;
            this.txtTenName.TextChanged += new System.EventHandler(this.txtTenName_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.Location = new System.Drawing.Point(51, 79);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(49, 14);
            this.label5.TabIndex = 8;
            this.label5.Text = "部数：";
            // 
            // txtBusu
            // 
            this.txtBusu.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtBusu.Location = new System.Drawing.Point(100, 74);
            this.txtBusu.Name = "txtBusu";
            this.txtBusu.Size = new System.Drawing.Size(76, 22);
            this.txtBusu.TabIndex = 9;
            this.txtBusu.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.Location = new System.Drawing.Point(9, 113);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(91, 14);
            this.label6.TabIndex = 10;
            this.label6.Text = "登録票番号：";
            // 
            // txtSNum
            // 
            this.txtSNum.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtSNum.Location = new System.Drawing.Point(100, 108);
            this.txtSNum.Name = "txtSNum";
            this.txtSNum.Size = new System.Drawing.Size(112, 22);
            this.txtSNum.TabIndex = 11;
            this.txtSNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtENum
            // 
            this.txtENum.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtENum.Location = new System.Drawing.Point(239, 108);
            this.txtENum.Name = "txtENum";
            this.txtENum.Size = new System.Drawing.Size(112, 22);
            this.txtENum.TabIndex = 12;
            this.txtENum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label7.Location = new System.Drawing.Point(215, 113);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(21, 14);
            this.label7.TabIndex = 13;
            this.label7.Text = "～";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label8.Location = new System.Drawing.Point(51, 147);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(49, 14);
            this.label8.TabIndex = 14;
            this.label8.Text = "代金：";
            // 
            // txtDaikin
            // 
            this.txtDaikin.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtDaikin.Location = new System.Drawing.Point(100, 142);
            this.txtDaikin.Name = "txtDaikin";
            this.txtDaikin.Size = new System.Drawing.Size(112, 22);
            this.txtDaikin.TabIndex = 15;
            this.txtDaikin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtDenNum);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.txtDaikin);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.txtENum);
            this.groupBox1.Controls.Add(this.txtSNum);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.txtBusu);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.txtTenName);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Location = new System.Drawing.Point(19, 411);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(887, 177);
            this.groupBox1.TabIndex = 16;
            this.groupBox1.TabStop = false;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(613, 600);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(192, 35);
            this.button1.TabIndex = 17;
            this.button1.Text = "納品書・請求書発行";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(811, 600);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(95, 35);
            this.button2.TabIndex = 18;
            this.button2.Text = "終了";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("ＭＳ ゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button3.Location = new System.Drawing.Point(841, 10);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(65, 28);
            this.button3.TabIndex = 19;
            this.button3.Text = "参照...";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label1
            // 
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(19, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(816, 28);
            this.label1.TabIndex = 20;
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtDenNum
            // 
            this.txtDenNum.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtDenNum.Location = new System.Drawing.Point(100, 21);
            this.txtDenNum.Name = "txtDenNum";
            this.txtDenNum.Size = new System.Drawing.Size(76, 22);
            this.txtDenNum.TabIndex = 17;
            this.txtDenNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(37, 26);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(63, 14);
            this.label2.TabIndex = 16;
            this.label2.Text = "伝票№：";
            // 
            // frmNouhinRep
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(918, 647);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.dataGridView1);
            this.Font = new System.Drawing.Font("ＭＳ ゴシック", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "frmNouhinRep";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "納品書・請求書発行";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmNouhinRep_FormClosing);
            this.Load += new System.EventHandler(this.frmNouhinRep_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtTenName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtBusu;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtSNum;
        private System.Windows.Forms.TextBox txtENum;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtDaikin;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtDenNum;
        private System.Windows.Forms.Label label2;
    }
}