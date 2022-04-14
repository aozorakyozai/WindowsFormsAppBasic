using System;

namespace WindowsFormsAppBasic
{
    partial class setting
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
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
            this.L1 = new System.Windows.Forms.Button();
            this.R1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // L1
            // 
            this.L1.Location = new System.Drawing.Point(1, 1);
            this.L1.Name = "L1";
            this.L1.Size = new System.Drawing.Size(80, 80);
            this.L1.TabIndex = 0;
            this.L1.Text = "<";
            this.L1.UseVisualStyleBackColor = true;
            this.L1.Click += new System.EventHandler(this.L1_Click);
            // 
            // R1
            // 
            this.R1.Location = new System.Drawing.Point(80, 1);
            this.R1.Name = "R1";
            this.R1.Size = new System.Drawing.Size(80, 80);
            this.R1.TabIndex = 1;
            this.R1.Text = ">";
            this.R1.UseVisualStyleBackColor = true;
            this.R1.Click += new System.EventHandler(this.R1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(159, 1);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(512, 80);
            this.textBox1.TabIndex = 2;
            // 
            // setting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ClientSize = new System.Drawing.Size(1017, 83);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.R1);
            this.Controls.Add(this.L1);
            this.Name = "setting";
            this.Text = "setting";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button L1;
        private System.Windows.Forms.Button R1;
        private System.Windows.Forms.TextBox textBox1;
        
    }
}

