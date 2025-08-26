namespace YYTools
{
    partial class SettingsForm
    {
        private System.ComponentModel.IContainer components = null;
        
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }
        
        private void InitializeComponent()
        {
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkAutoScale = new System.Windows.Forms.CheckBox();
            this.numFontSize = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkEnableDebugLog = new System.Windows.Forms.CheckBox();
            this.chkWPSPriority = new System.Windows.Forms.CheckBox();
            this.cmbPerformanceMode = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.txtBillName = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtBillProduct = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtBillTrack = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.txtShippingName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtShippingProduct = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtShippingTrack = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.btnBrowseLog = new System.Windows.Forms.Button();
            this.txtLogDirectory = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.numProgressFreq = new System.Windows.Forms.NumericUpDown();
            this.label6 = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnApply = new System.Windows.Forms.Button();
            this.btnResetDefaults = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numFontSize)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.groupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numProgressFreq)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(460, 320);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.groupBox2);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(452, 294);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "基本设置";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkAutoScale);
            this.groupBox1.Controls.Add(this.numFontSize);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(16, 16);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(420, 80);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "界面设置";
            // 
            // chkAutoScale
            // 
            this.chkAutoScale.AutoSize = true;
            this.chkAutoScale.Location = new System.Drawing.Point(20, 50);
            this.chkAutoScale.Name = "chkAutoScale";
            this.chkAutoScale.Size = new System.Drawing.Size(96, 16);
            this.chkAutoScale.TabIndex = 2;
            this.chkAutoScale.Text = "自动DPI缩放";
            this.chkAutoScale.UseVisualStyleBackColor = true;
            // 
            // numFontSize
            // 
            this.numFontSize.Location = new System.Drawing.Point(80, 20);
            this.numFontSize.Maximum = new decimal(new int[] {
            16,
            0,
            0,
            0});
            this.numFontSize.Minimum = new decimal(new int[] {
            8,
            0,
            0,
            0});
            this.numFontSize.Name = "numFontSize";
            this.numFontSize.Size = new System.Drawing.Size(60, 21);
            this.numFontSize.TabIndex = 1;
            this.numFontSize.Value = new decimal(new int[] {
            9,
            0,
            0,
            0});
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "字体大小:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkEnableDebugLog);
            this.groupBox2.Controls.Add(this.chkWPSPriority);
            this.groupBox2.Controls.Add(this.cmbPerformanceMode);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Location = new System.Drawing.Point(16, 110);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(420, 120);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "性能和兼容性";
            // 
            // chkEnableDebugLog
            // 
            this.chkEnableDebugLog.AutoSize = true;
            this.chkEnableDebugLog.Location = new System.Drawing.Point(20, 84);
            this.chkEnableDebugLog.Name = "chkEnableDebugLog";
            this.chkEnableDebugLog.Size = new System.Drawing.Size(96, 16);
            this.chkEnableDebugLog.TabIndex = 3;
            this.chkEnableDebugLog.Text = "启用调试日志";
            this.chkEnableDebugLog.UseVisualStyleBackColor = true;
            // 
            // chkWPSPriority
            // 
            this.chkWPSPriority.AutoSize = true;
            this.chkWPSPriority.Location = new System.Drawing.Point(20, 54);
            this.chkWPSPriority.Name = "chkWPSPriority";
            this.chkWPSPriority.Size = new System.Drawing.Size(144, 16);
            this.chkWPSPriority.TabIndex = 2;
            this.chkWPSPriority.Text = "WPS表格优先（推荐）";
            this.chkWPSPriority.UseVisualStyleBackColor = true;
            // 
            // cmbPerformanceMode
            // 
            this.cmbPerformanceMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPerformanceMode.FormattingEnabled = true;
            this.cmbPerformanceMode.Items.AddRange(new object[] {
            "极速模式（推荐）",
            "平衡模式",
            "兼容模式"});
            this.cmbPerformanceMode.Location = new System.Drawing.Point(80, 20);
            this.cmbPerformanceMode.Name = "cmbPerformanceMode";
            this.cmbPerformanceMode.Size = new System.Drawing.Size(140, 20);
            this.cmbPerformanceMode.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 0;
            this.label2.Text = "性能模式:";
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.groupBox4);
            this.tabPage2.Controls.Add(this.groupBox3);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(452, 294);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "默认列设置";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.txtBillName);
            this.groupBox4.Controls.Add(this.label8);
            this.groupBox4.Controls.Add(this.txtBillProduct);
            this.groupBox4.Controls.Add(this.label9);
            this.groupBox4.Controls.Add(this.txtBillTrack);
            this.groupBox4.Controls.Add(this.label10);
            this.groupBox4.Location = new System.Drawing.Point(16, 160);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(420, 120);
            this.groupBox4.TabIndex = 1;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "账单明细";
            // 
            // txtBillName
            // 
            this.txtBillName.Location = new System.Drawing.Point(100, 80);
            this.txtBillName.Name = "txtBillName";
            this.txtBillName.Size = new System.Drawing.Size(60, 21);
            this.txtBillName.TabIndex = 5;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(20, 83);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(65, 12);
            this.label8.TabIndex = 4;
            this.label8.Text = "商品名称列:";
            // 
            // txtBillProduct
            // 
            this.txtBillProduct.Location = new System.Drawing.Point(100, 50);
            this.txtBillProduct.Name = "txtBillProduct";
            this.txtBillProduct.Size = new System.Drawing.Size(60, 21);
            this.txtBillProduct.TabIndex = 3;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(20, 53);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(65, 12);
            this.label9.TabIndex = 2;
            this.label9.Text = "商品编码列:";
            // 
            // txtBillTrack
            // 
            this.txtBillTrack.Location = new System.Drawing.Point(100, 20);
            this.txtBillTrack.Name = "txtBillTrack";
            this.txtBillTrack.Size = new System.Drawing.Size(60, 21);
            this.txtBillTrack.TabIndex = 1;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(20, 23);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(53, 12);
            this.label10.TabIndex = 0;
            this.label10.Text = "运单号列:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.txtShippingName);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.txtShippingProduct);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.txtShippingTrack);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Location = new System.Drawing.Point(16, 16);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(420, 120);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "发货明细";
            // 
            // txtShippingName
            // 
            this.txtShippingName.Location = new System.Drawing.Point(100, 80);
            this.txtShippingName.Name = "txtShippingName";
            this.txtShippingName.Size = new System.Drawing.Size(60, 21);
            this.txtShippingName.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(20, 83);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 12);
            this.label5.TabIndex = 4;
            this.label5.Text = "商品名称列:";
            // 
            // txtShippingProduct
            // 
            this.txtShippingProduct.Location = new System.Drawing.Point(100, 50);
            this.txtShippingProduct.Name = "txtShippingProduct";
            this.txtShippingProduct.Size = new System.Drawing.Size(60, 21);
            this.txtShippingProduct.TabIndex = 3;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(20, 53);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 2;
            this.label4.Text = "商品编码列:";
            // 
            // txtShippingTrack
            // 
            this.txtShippingTrack.Location = new System.Drawing.Point(100, 20);
            this.txtShippingTrack.Name = "txtShippingTrack";
            this.txtShippingTrack.Size = new System.Drawing.Size(60, 21);
            this.txtShippingTrack.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(20, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 0;
            this.label3.Text = "运单号列:";
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.groupBox5);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(452, 294);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "高级设置";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.btnBrowseLog);
            this.groupBox5.Controls.Add(this.txtLogDirectory);
            this.groupBox5.Controls.Add(this.label7);
            this.groupBox5.Controls.Add(this.numProgressFreq);
            this.groupBox5.Controls.Add(this.label6);
            this.groupBox5.Location = new System.Drawing.Point(16, 16);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(420, 120);
            this.groupBox5.TabIndex = 0;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "高级选项";
            // 
            // btnBrowseLog
            // 
            this.btnBrowseLog.Location = new System.Drawing.Point(330, 50);
            this.btnBrowseLog.Name = "btnBrowseLog";
            this.btnBrowseLog.Size = new System.Drawing.Size(60, 23);
            this.btnBrowseLog.TabIndex = 4;
            this.btnBrowseLog.Text = "浏览...";
            this.btnBrowseLog.UseVisualStyleBackColor = true;
            this.btnBrowseLog.Click += new System.EventHandler(this.btnBrowseLog_Click);
            // 
            // txtLogDirectory
            // 
            this.txtLogDirectory.Location = new System.Drawing.Point(100, 52);
            this.txtLogDirectory.Name = "txtLogDirectory";
            this.txtLogDirectory.Size = new System.Drawing.Size(220, 21);
            this.txtLogDirectory.TabIndex = 3;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(20, 55);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 12);
            this.label7.TabIndex = 2;
            this.label7.Text = "日志目录:";
            // 
            // numProgressFreq
            // 
            this.numProgressFreq.Location = new System.Drawing.Point(100, 20);
            this.numProgressFreq.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numProgressFreq.Minimum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numProgressFreq.Name = "numProgressFreq";
            this.numProgressFreq.Size = new System.Drawing.Size(80, 21);
            this.numProgressFreq.TabIndex = 1;
            this.numProgressFreq.Value = new decimal(new int[] {
            100,
            0,
            0,
            0});
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(20, 22);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(77, 12);
            this.label6.TabIndex = 0;
            this.label6.Text = "进度更新频率:";
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(235, 350);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 28);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "确定";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(320, 350);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 28);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnApply
            // 
            this.btnApply.Location = new System.Drawing.Point(405, 350);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(75, 28);
            this.btnApply.TabIndex = 3;
            this.btnApply.Text = "应用";
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.btnApply_Click);
            // 
            // btnResetDefaults
            // 
            this.btnResetDefaults.Location = new System.Drawing.Point(20, 350);
            this.btnResetDefaults.Name = "btnResetDefaults";
            this.btnResetDefaults.Size = new System.Drawing.Size(90, 28);
            this.btnResetDefaults.TabIndex = 4;
            this.btnResetDefaults.Text = "恢复默认";
            this.btnResetDefaults.UseVisualStyleBackColor = true;
            this.btnResetDefaults.Click += new System.EventHandler(this.btnResetDefaults_Click);
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(494, 392);
            this.Controls.Add(this.btnResetDefaults);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "YY工具设置 - WPS优先";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numFontSize)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numProgressFreq)).EndInit();
            this.ResumeLayout(false);
        }
        
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown numFontSize;
        private System.Windows.Forms.CheckBox chkAutoScale;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbPerformanceMode;
        private System.Windows.Forms.CheckBox chkWPSPriority;
        private System.Windows.Forms.CheckBox chkEnableDebugLog;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtShippingTrack;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtShippingProduct;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtShippingName;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox txtBillTrack;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtBillProduct;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtBillName;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.NumericUpDown numProgressFreq;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtLogDirectory;
        private System.Windows.Forms.Button btnBrowseLog;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.Button btnResetDefaults;
    }
}
