namespace PlusTest
{
    partial class fmVolTop
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
            this.m_VolTopGrid = new System.Windows.Forms.DataGridView();
            this.m_rdoV1 = new System.Windows.Forms.RadioButton();
            this.m_rdoV2 = new System.Windows.Forms.RadioButton();
            this.m_chkExc1 = new System.Windows.Forms.CheckBox();
            this.m_chkExc2 = new System.Windows.Forms.CheckBox();
            this.m_rdoMarketAll = new System.Windows.Forms.RadioButton();
            this.m_rdoMarket1 = new System.Windows.Forms.RadioButton();
            this.m_rdoMarket2 = new System.Windows.Forms.RadioButton();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.m_btnRetry = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.m_VolTopGrid)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // m_VolTopGrid
            // 
            this.m_VolTopGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_VolTopGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.m_VolTopGrid.Location = new System.Drawing.Point(12, 43);
            this.m_VolTopGrid.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.m_VolTopGrid.Name = "m_VolTopGrid";
            this.m_VolTopGrid.RowTemplate.Height = 23;
            this.m_VolTopGrid.Size = new System.Drawing.Size(848, 455);
            this.m_VolTopGrid.TabIndex = 0;
            // 
            // m_rdoV1
            // 
            this.m_rdoV1.AutoSize = true;
            this.m_rdoV1.Checked = true;
            this.m_rdoV1.Location = new System.Drawing.Point(5, 6);
            this.m_rdoV1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.m_rdoV1.Name = "m_rdoV1";
            this.m_rdoV1.Size = new System.Drawing.Size(96, 21);
            this.m_rdoV1.TabIndex = 1;
            this.m_rdoV1.TabStop = true;
            this.m_rdoV1.Text = "거래량 상위";
            this.m_rdoV1.UseVisualStyleBackColor = true;
            this.m_rdoV1.Click += new System.EventHandler(this.m_rdoV1_Click);
            // 
            // m_rdoV2
            // 
            this.m_rdoV2.AutoSize = true;
            this.m_rdoV2.Location = new System.Drawing.Point(103, 6);
            this.m_rdoV2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.m_rdoV2.Name = "m_rdoV2";
            this.m_rdoV2.Size = new System.Drawing.Size(109, 21);
            this.m_rdoV2.TabIndex = 2;
            this.m_rdoV2.Text = "거래대금 상위";
            this.m_rdoV2.UseVisualStyleBackColor = true;
            this.m_rdoV2.Click += new System.EventHandler(this.m_rdoV2_Click);
            // 
            // m_chkExc1
            // 
            this.m_chkExc1.AutoSize = true;
            this.m_chkExc1.Location = new System.Drawing.Point(510, 16);
            this.m_chkExc1.Name = "m_chkExc1";
            this.m_chkExc1.Size = new System.Drawing.Size(79, 21);
            this.m_chkExc1.TabIndex = 3;
            this.m_chkExc1.Text = "관리제외";
            this.m_chkExc1.UseVisualStyleBackColor = true;
            this.m_chkExc1.Click += new System.EventHandler(this.m_chkExc1_Click);
            // 
            // m_chkExc2
            // 
            this.m_chkExc2.AutoSize = true;
            this.m_chkExc2.Location = new System.Drawing.Point(600, 15);
            this.m_chkExc2.Name = "m_chkExc2";
            this.m_chkExc2.Size = new System.Drawing.Size(92, 21);
            this.m_chkExc2.TabIndex = 3;
            this.m_chkExc2.Text = "우선주제외";
            this.m_chkExc2.UseVisualStyleBackColor = true;
            this.m_chkExc2.Click += new System.EventHandler(this.m_chkExc2_Click);
            // 
            // m_rdoMarketAll
            // 
            this.m_rdoMarketAll.AutoSize = true;
            this.m_rdoMarketAll.Checked = true;
            this.m_rdoMarketAll.Location = new System.Drawing.Point(14, 6);
            this.m_rdoMarketAll.Name = "m_rdoMarketAll";
            this.m_rdoMarketAll.Size = new System.Drawing.Size(52, 21);
            this.m_rdoMarketAll.TabIndex = 4;
            this.m_rdoMarketAll.TabStop = true;
            this.m_rdoMarketAll.Text = "전체";
            this.m_rdoMarketAll.UseVisualStyleBackColor = true;
            this.m_rdoMarketAll.Click += new System.EventHandler(this.m_rdoMarketAll_Click);
            // 
            // m_rdoMarket1
            // 
            this.m_rdoMarket1.AutoSize = true;
            this.m_rdoMarket1.Location = new System.Drawing.Point(66, 6);
            this.m_rdoMarket1.Name = "m_rdoMarket1";
            this.m_rdoMarket1.Size = new System.Drawing.Size(65, 21);
            this.m_rdoMarket1.TabIndex = 4;
            this.m_rdoMarket1.Text = "거래소";
            this.m_rdoMarket1.UseVisualStyleBackColor = true;
            this.m_rdoMarket1.Click += new System.EventHandler(this.m_rdoMarket1_Click);
            // 
            // m_rdoMarket2
            // 
            this.m_rdoMarket2.AutoSize = true;
            this.m_rdoMarket2.Location = new System.Drawing.Point(132, 6);
            this.m_rdoMarket2.Name = "m_rdoMarket2";
            this.m_rdoMarket2.Size = new System.Drawing.Size(65, 21);
            this.m_rdoMarket2.TabIndex = 4;
            this.m_rdoMarket2.Text = "코스닥";
            this.m_rdoMarket2.UseVisualStyleBackColor = true;
            this.m_rdoMarket2.Click += new System.EventHandler(this.m_rdoMarket2_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.m_rdoMarketAll);
            this.panel1.Controls.Add(this.m_rdoMarket2);
            this.panel1.Controls.Add(this.m_rdoMarket1);
            this.panel1.Location = new System.Drawing.Point(12, 10);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(213, 30);
            this.panel1.TabIndex = 5;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.m_rdoV1);
            this.panel2.Controls.Add(this.m_rdoV2);
            this.panel2.Location = new System.Drawing.Point(232, 10);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(216, 30);
            this.panel2.TabIndex = 6;
            // 
            // m_btnRetry
            // 
            this.m_btnRetry.Location = new System.Drawing.Point(787, 10);
            this.m_btnRetry.Name = "m_btnRetry";
            this.m_btnRetry.Size = new System.Drawing.Size(75, 23);
            this.m_btnRetry.TabIndex = 7;
            this.m_btnRetry.Text = "조회";
            this.m_btnRetry.UseVisualStyleBackColor = true;
            this.m_btnRetry.Click += new System.EventHandler(this.m_btnRetry_Click);
            // 
            // fmVolTop
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(872, 509);
            this.Controls.Add(this.m_btnRetry);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.m_chkExc2);
            this.Controls.Add(this.m_chkExc1);
            this.Controls.Add(this.m_VolTopGrid);
            this.Font = new System.Drawing.Font("Malgun Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "fmVolTop";
            this.Text = "거래량/거래대금 상위";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.m_VolTopGrid)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView m_VolTopGrid;
        private System.Windows.Forms.RadioButton m_rdoV1;
        private System.Windows.Forms.RadioButton m_rdoV2;
        private System.Windows.Forms.CheckBox m_chkExc1;
        private System.Windows.Forms.CheckBox m_chkExc2;
        private System.Windows.Forms.RadioButton m_rdoMarketAll;
        private System.Windows.Forms.RadioButton m_rdoMarket1;
        private System.Windows.Forms.RadioButton m_rdoMarket2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button m_btnRetry;
    }
}