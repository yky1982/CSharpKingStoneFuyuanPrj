namespace KingStoneFuYuan
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.Grp_No2 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.Grp_No1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.Grp_No2.SuspendLayout();
            this.Grp_No1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Grp_No2
            // 
            this.Grp_No2.Controls.Add(this.label2);
            this.Grp_No2.Location = new System.Drawing.Point(12, 498);
            this.Grp_No2.Name = "Grp_No2";
            this.Grp_No2.Size = new System.Drawing.Size(1240, 475);
            this.Grp_No2.TabIndex = 4;
            this.Grp_No2.TabStop = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label2.Location = new System.Drawing.Point(588, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(122, 21);
            this.label2.TabIndex = 1;
            this.label2.Text = "2#试验装置";
            // 
            // Grp_No1
            // 
            this.Grp_No1.Controls.Add(this.label1);
            this.Grp_No1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Grp_No1.Location = new System.Drawing.Point(12, 13);
            this.Grp_No1.Name = "Grp_No1";
            this.Grp_No1.Size = new System.Drawing.Size(1240, 475);
            this.Grp_No1.TabIndex = 3;
            this.Grp_No1.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label1.Location = new System.Drawing.Point(588, 3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(122, 21);
            this.label1.TabIndex = 0;
            this.label1.Text = "1#试验装置";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1264, 986);
            this.Controls.Add(this.Grp_No2);
            this.Controls.Add(this.Grp_No1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Grp_No2.ResumeLayout(false);
            this.Grp_No2.PerformLayout();
            this.Grp_No1.ResumeLayout(false);
            this.Grp_No1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox Grp_No2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox Grp_No1;
        private System.Windows.Forms.Label label1;
    }
}

