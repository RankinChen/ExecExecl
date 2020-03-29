namespace ExecExecl
{
    partial class MainForm
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.gBoxInput = new System.Windows.Forms.GroupBox();
            this.btnBrowseSourceFile = new System.Windows.Forms.Button();
            this.txtSourceFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.ofdBrowserFile = new System.Windows.Forms.OpenFileDialog();
            this.sfdSaveNewFile = new System.Windows.Forms.SaveFileDialog();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.ltbErrorMsg = new System.Windows.Forms.ListBox();
            this.btnExec = new System.Windows.Forms.Button();
            this.gBoxInput.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // gBoxInput
            // 
            this.gBoxInput.Controls.Add(this.btnBrowseSourceFile);
            this.gBoxInput.Controls.Add(this.txtSourceFile);
            this.gBoxInput.Controls.Add(this.label1);
            this.gBoxInput.Location = new System.Drawing.Point(12, 12);
            this.gBoxInput.Name = "gBoxInput";
            this.gBoxInput.Size = new System.Drawing.Size(776, 63);
            this.gBoxInput.TabIndex = 1;
            this.gBoxInput.TabStop = false;
            this.gBoxInput.Text = "选择处理文件";
            // 
            // btnBrowseSourceFile
            // 
            this.btnBrowseSourceFile.Location = new System.Drawing.Point(672, 23);
            this.btnBrowseSourceFile.Name = "btnBrowseSourceFile";
            this.btnBrowseSourceFile.Size = new System.Drawing.Size(98, 23);
            this.btnBrowseSourceFile.TabIndex = 2;
            this.btnBrowseSourceFile.Text = "选择原始文件";
            this.btnBrowseSourceFile.UseVisualStyleBackColor = true;
            this.btnBrowseSourceFile.Click += new System.EventHandler(this.btnBrowseSourceFile_Click);
            // 
            // txtSourceFile
            // 
            this.txtSourceFile.Location = new System.Drawing.Point(77, 25);
            this.txtSourceFile.Name = "txtSourceFile";
            this.txtSourceFile.Size = new System.Drawing.Size(589, 21);
            this.txtSourceFile.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "原始Excel:";
            // 
            // ofdBrowserFile
            // 
            this.ofdBrowserFile.Title = "原始Excel";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.ltbErrorMsg);
            this.groupBox5.Location = new System.Drawing.Point(12, 159);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(776, 289);
            this.groupBox5.TabIndex = 11;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "日志记录";
            // 
            // ltbErrorMsg
            // 
            this.ltbErrorMsg.FormattingEnabled = true;
            this.ltbErrorMsg.ItemHeight = 12;
            this.ltbErrorMsg.Location = new System.Drawing.Point(6, 24);
            this.ltbErrorMsg.Name = "ltbErrorMsg";
            this.ltbErrorMsg.Size = new System.Drawing.Size(764, 184);
            this.ltbErrorMsg.TabIndex = 0;
            // 
            // btnExec
            // 
            this.btnExec.Location = new System.Drawing.Point(268, 90);
            this.btnExec.Name = "btnExec";
            this.btnExec.Size = new System.Drawing.Size(75, 23);
            this.btnExec.TabIndex = 12;
            this.btnExec.Text = "开始处理";
            this.btnExec.UseVisualStyleBackColor = true;
            this.btnExec.Click += new System.EventHandler(this.btnExec_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnExec);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.gBoxInput);
            this.Name = "MainForm";
            this.Text = "MainForm";
            this.gBoxInput.ResumeLayout(false);
            this.gBoxInput.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gBoxInput;
        private System.Windows.Forms.Button btnBrowseSourceFile;
        private System.Windows.Forms.TextBox txtSourceFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog ofdBrowserFile;
        private System.Windows.Forms.SaveFileDialog sfdSaveNewFile;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.ListBox ltbErrorMsg;
        private System.Windows.Forms.Button btnExec;
    }
}

