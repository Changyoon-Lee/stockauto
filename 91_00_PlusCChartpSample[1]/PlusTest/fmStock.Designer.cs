namespace PlusTest
{
    partial class fmStock
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
            this.m_HogaView = new System.Windows.Forms.DataGridView();
            this.m_etCode = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.m_HogaView)).BeginInit();
            this.SuspendLayout();
            // 
            // m_HogaView
            // 
            this.m_HogaView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.m_HogaView.Location = new System.Drawing.Point(7, 41);
            this.m_HogaView.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.m_HogaView.Name = "m_HogaView";
            this.m_HogaView.RowTemplate.Height = 23;
            this.m_HogaView.Size = new System.Drawing.Size(416, 557);
            this.m_HogaView.TabIndex = 0;
            // 
            // m_etCode
            // 
            this.m_etCode.Location = new System.Drawing.Point(8, 11);
            this.m_etCode.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.m_etCode.MaxLength = 6;
            this.m_etCode.Name = "m_etCode";
            this.m_etCode.Size = new System.Drawing.Size(119, 25);
            this.m_etCode.TabIndex = 1;
            this.m_etCode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.m_etCode_KeyDown);
            // 
            // fmStock
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(432, 611);
            this.Controls.Add(this.m_etCode);
            this.Controls.Add(this.m_HogaView);
            this.Font = new System.Drawing.Font("Malgun Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "fmStock";
            this.Text = "주식현재가";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.fmStock_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.m_HogaView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView m_HogaView;
        private System.Windows.Forms.TextBox m_etCode;
    }
}

