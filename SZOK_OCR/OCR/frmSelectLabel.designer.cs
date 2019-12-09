namespace SZOK_OCR.OCR
{
    partial class frmSelectLabel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSelectLabel));
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.Dgv1 = new System.Windows.Forms.DataGridView();
            this.lblDataCnt = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.Dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.Control;
            this.button1.Font = new System.Drawing.Font("ＭＳ ゴシック", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button1.Location = new System.Drawing.Point(272, 478);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(78, 28);
            this.button1.TabIndex = 4;
            this.button1.Text = "決定";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.SystemColors.Control;
            this.button2.Font = new System.Drawing.Font("ＭＳ ゴシック", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button2.Location = new System.Drawing.Point(355, 478);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(78, 28);
            this.button2.TabIndex = 5;
            this.button2.Text = "中止";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Dgv1
            // 
            this.Dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Dgv1.Location = new System.Drawing.Point(12, 11);
            this.Dgv1.Name = "Dgv1";
            this.Dgv1.RowTemplate.Height = 21;
            this.Dgv1.Size = new System.Drawing.Size(421, 402);
            this.Dgv1.TabIndex = 7;
            this.Dgv1.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.Dgv1_CellValueChanged);
            this.Dgv1.CurrentCellDirtyStateChanged += new System.EventHandler(this.Dgv1_CurrentCellDirtyStateChanged);
            // 
            // lblDataCnt
            // 
            this.lblDataCnt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblDataCnt.Font = new System.Drawing.Font("ＭＳ ゴシック", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblDataCnt.Location = new System.Drawing.Point(376, 427);
            this.lblDataCnt.Name = "lblDataCnt";
            this.lblDataCnt.Size = new System.Drawing.Size(57, 26);
            this.lblDataCnt.TabIndex = 8;
            this.lblDataCnt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("ＭＳ ゴシック", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(139, 432);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(232, 16);
            this.label1.TabIndex = 9;
            this.label1.Text = "現在、処理中の防犯登録データ\r\n";
            // 
            // frmFaxSelectHaken
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(444, 518);
            this.ControlBox = false;
            this.Controls.Add(this.lblDataCnt);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Dgv1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Font = new System.Drawing.Font("ＭＳ ゴシック", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmFaxSelectHaken";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "処理を行うラベルを選択してください";
            this.Load += new System.EventHandler(this.frmFaxSelect_Load);
            ((System.ComponentModel.ISupportInitialize)(this.Dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridView Dgv1;
        private System.Windows.Forms.Label lblDataCnt;
        private System.Windows.Forms.Label label1;
    }
}