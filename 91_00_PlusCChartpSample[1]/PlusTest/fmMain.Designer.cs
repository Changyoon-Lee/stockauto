namespace PlusTest
{
    partial class fmMain
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
            this.m_btnForm1 = new System.Windows.Forms.Button();
            this.m_btnForm2 = new System.Windows.Forms.Button();
            this.m_btnForm3 = new System.Windows.Forms.Button();
            this.m_btnForm4 = new System.Windows.Forms.Button();
            this.m_btnForm5 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // m_btnForm1
            // 
            this.m_btnForm1.Location = new System.Drawing.Point(38, 35);
            this.m_btnForm1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.m_btnForm1.Name = "m_btnForm1";
            this.m_btnForm1.Size = new System.Drawing.Size(136, 33);
            this.m_btnForm1.TabIndex = 0;
            this.m_btnForm1.Text = "주식현재가";
            this.m_btnForm1.UseVisualStyleBackColor = true;
            this.m_btnForm1.Click += new System.EventHandler(this.m_btnForm1_Click);
            // 
            // m_btnForm2
            // 
            this.m_btnForm2.Location = new System.Drawing.Point(38, 76);
            this.m_btnForm2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.m_btnForm2.Name = "m_btnForm2";
            this.m_btnForm2.Size = new System.Drawing.Size(136, 33);
            this.m_btnForm2.TabIndex = 1;
            this.m_btnForm2.Text = "선물현재가";
            this.m_btnForm2.UseVisualStyleBackColor = true;
            this.m_btnForm2.Click += new System.EventHandler(this.m_btnForm2_Click);
            // 
            // m_btnForm3
            // 
            this.m_btnForm3.Location = new System.Drawing.Point(38, 117);
            this.m_btnForm3.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.m_btnForm3.Name = "m_btnForm3";
            this.m_btnForm3.Size = new System.Drawing.Size(136, 33);
            this.m_btnForm3.TabIndex = 2;
            this.m_btnForm3.Text = "옵션현재가";
            this.m_btnForm3.UseVisualStyleBackColor = true;
            this.m_btnForm3.Click += new System.EventHandler(this.m_btnForm3_Click);
            // 
            // m_btnForm4
            // 
            this.m_btnForm4.Location = new System.Drawing.Point(38, 158);
            this.m_btnForm4.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.m_btnForm4.Name = "m_btnForm4";
            this.m_btnForm4.Size = new System.Drawing.Size(136, 33);
            this.m_btnForm4.TabIndex = 3;
            this.m_btnForm4.Text = "주식선물현재가";
            this.m_btnForm4.UseVisualStyleBackColor = true;
            this.m_btnForm4.Click += new System.EventHandler(this.m_btnForm4_Click);
            // 
            // m_btnForm5
            // 
            this.m_btnForm5.Location = new System.Drawing.Point(38, 199);
            this.m_btnForm5.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.m_btnForm5.Name = "m_btnForm5";
            this.m_btnForm5.Size = new System.Drawing.Size(136, 33);
            this.m_btnForm5.TabIndex = 3;
            this.m_btnForm5.Text = "거래량상위종목";
            this.m_btnForm5.UseVisualStyleBackColor = true;
            this.m_btnForm5.Click += new System.EventHandler(this.m_btnForm5_Click);
            // 
            // fmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(234, 637);
            this.Controls.Add(this.m_btnForm5);
            this.Controls.Add(this.m_btnForm4);
            this.Controls.Add(this.m_btnForm3);
            this.Controls.Add(this.m_btnForm2);
            this.Controls.Add(this.m_btnForm1);
            this.Font = new System.Drawing.Font("Malgun Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "fmMain";
            this.Text = "Plus 샘플 자료";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button m_btnForm1;
        private System.Windows.Forms.Button m_btnForm2;
        private System.Windows.Forms.Button m_btnForm3;
        private System.Windows.Forms.Button m_btnForm4;
        private System.Windows.Forms.Button m_btnForm5;
    }
}