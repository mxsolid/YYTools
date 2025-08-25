namespace YYTools
{
    partial class SettingsForm
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
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabPerformance = new System.Windows.Forms.TabPage();
            this.tabUI = new System.Windows.Forms.TabPage();
            this.tabDefaults = new System.Windows.Forms.TabPage();
            this.tabAdvanced = new System.Windows.Forms.TabPage();
            this.grpPerformance = new System.Windows.Forms.GroupBox();
            this.lblPerformanceMode = new System.Windows.Forms.Label();
            this.cmbPerformanceMode = new System.Windows.Forms.ComboBox();
            this.lblPerformanceDesc = new System.Windows.Forms.Label();
            this.grpUI = new System.Windows.Forms.GroupBox();
            this.lblFontSize = new System.Windows.Forms.Label();
            this.nudFontSize = new System.Windows.Forms.NumericUpDown();
            this.chkAutoScale = new System.Windows.Forms.CheckBox();
            this.grpShippingDefaults = new System.Windows.Forms.GroupBox();
            this.lblDefaultShippingTrack = new System.Windows.Forms.Label();
            this.txtDefaultShippingTrack = new System.Windows.Forms.TextBox();
            this.lblDefaultShippingProduct = new System.Windows.Forms.Label();
            this.txtDefaultShippingProduct = new System.Windows.Forms.TextBox();
            this.lblDefaultShippingName = new System.Windows.Forms.Label();
            this.txtDefaultShippingName = new System.Windows.Forms.TextBox();
            this.grpBillDefaults = new System.Windows.Forms.GroupBox();
            this.lblDefaultBillTrack = new System.Windows.Forms.Label();
            this.txtDefaultBillTrack = new System.Windows.Forms.TextBox();
            this.lblDefaultBillProduct = new System.Windows.Forms.Label();
            this.txtDefaultBillProduct = new System.Windows.Forms.TextBox();
            this.lblDefaultBillName = new System.Windows.Forms.Label();
            this.txtDefaultBillName = new System.Windows.Forms.TextBox();
            this.grpAdvanced = new System.Windows.Forms.GroupBox();
            this.lblLogDirectory = new System.Windows.Forms.Label();
            this.txtLogDirectory = new System.Windows.Forms.TextBox();
            this.btnBrowseLog = new System.Windows.Forms.Button();
            this.chkAutoSelectSheets = new System.Windows.Forms.CheckBox();
            this.lblProgressUpdate = new System.Windows.Forms.Label();
            this.nudProgressUpdateInterval = new System.Windows.Forms.NumericUpDown();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.tabControl.SuspendLayout();
            this.tabPerformance.SuspendLayout();
            this.tabUI.SuspendLayout();
            this.tabDefaults.SuspendLayout();
            this.tabAdvanced.SuspendLayout();
            this.grpPerformance.SuspendLayout();
            this.grpUI.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudFontSize)).BeginInit();
            this.grpShippingDefaults.SuspendLayout();
            this.grpBillDefaults.SuspendLayout();
            this.grpAdvanced.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudProgressUpdateInterval)).BeginInit();
            this.SuspendLayout();
            
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.tabPerformance);
            this.tabControl.Controls.Add(this.tabUI);
            this.tabControl.Controls.Add(this.tabDefaults);
            this.tabControl.Controls.Add(this.tabAdvanced);
            this.tabControl.Location = new System.Drawing.Point(12, 12);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(560, 420);
            this.tabControl.TabIndex = 0;
            this.tabControl.Font = new System.Drawing.Font("微软雅黑", 9F);
            
            // 
            // tabPerformance - 性能设置
            // 
            this.tabPerformance.Controls.Add(this.grpPerformance);
            this.tabPerformance.Location = new System.Drawing.Point(4, 26);
            this.tabPerformance.Name = "tabPerformance";
            this.tabPerformance.Padding = new System.Windows.Forms.Padding(3);
            this.tabPerformance.Size = new System.Drawing.Size(552, 390);
            this.tabPerformance.TabIndex = 0;
            this.tabPerformance.Text = "⚡ 性能设置";
            this.tabPerformance.UseVisualStyleBackColor = true;
            
            this.grpPerformance.Controls.Add(this.lblPerformanceMode);
            this.grpPerformance.Controls.Add(this.cmbPerformanceMode);
            this.grpPerformance.Controls.Add(this.lblPerformanceDesc);
            this.grpPerformance.Location = new System.Drawing.Point(20, 20);
            this.grpPerformance.Name = "grpPerformance";
            this.grpPerformance.Size = new System.Drawing.Size(512, 120);
            this.grpPerformance.TabIndex = 0;
            this.grpPerformance.TabStop = false;
            this.grpPerformance.Text = "性能模式选择";
            this.grpPerformance.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            
            this.lblPerformanceMode.AutoSize = true;
            this.lblPerformanceMode.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblPerformanceMode.Location = new System.Drawing.Point(20, 35);
            this.lblPerformanceMode.Name = "lblPerformanceMode";
            this.lblPerformanceMode.Size = new System.Drawing.Size(68, 17);
            this.lblPerformanceMode.TabIndex = 0;
            this.lblPerformanceMode.Text = "性能模式：";
            
            this.cmbPerformanceMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPerformanceMode.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.cmbPerformanceMode.FormattingEnabled = true;
            this.cmbPerformanceMode.Items.AddRange(new object[] {
                "极速模式 (推荐)",
                "平衡模式",
                "兼容模式"});
            this.cmbPerformanceMode.Location = new System.Drawing.Point(100, 32);
            this.cmbPerformanceMode.Name = "cmbPerformanceMode";
            this.cmbPerformanceMode.Size = new System.Drawing.Size(200, 25);
            this.cmbPerformanceMode.TabIndex = 1;
            this.cmbPerformanceMode.SelectedIndexChanged += new System.EventHandler(this.cmbPerformanceMode_SelectedIndexChanged);
            
            this.lblPerformanceDesc.Font = new System.Drawing.Font("微软雅黑", 8.5F);
            this.lblPerformanceDesc.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lblPerformanceDesc.Location = new System.Drawing.Point(20, 70);
            this.lblPerformanceDesc.Name = "lblPerformanceDesc";
            this.lblPerformanceDesc.Size = new System.Drawing.Size(470, 35);
            this.lblPerformanceDesc.TabIndex = 2;
            this.lblPerformanceDesc.Text = "极速模式：最高性能，适用于高配置机器（推荐）";
            
            // 
            // tabUI - 界面设置
            // 
            this.tabUI.Controls.Add(this.grpUI);
            this.tabUI.Location = new System.Drawing.Point(4, 26);
            this.tabUI.Name = "tabUI";
            this.tabUI.Padding = new System.Windows.Forms.Padding(3);
            this.tabUI.Size = new System.Drawing.Size(552, 390);
            this.tabUI.TabIndex = 1;
            this.tabUI.Text = "🎨 界面设置";
            this.tabUI.UseVisualStyleBackColor = true;
            
            this.grpUI.Controls.Add(this.lblFontSize);
            this.grpUI.Controls.Add(this.nudFontSize);
            this.grpUI.Controls.Add(this.chkAutoScale);
            this.grpUI.Location = new System.Drawing.Point(20, 20);
            this.grpUI.Name = "grpUI";
            this.grpUI.Size = new System.Drawing.Size(512, 120);
            this.grpUI.TabIndex = 0;
            this.grpUI.TabStop = false;
            this.grpUI.Text = "字体和显示设置";
            this.grpUI.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            
            this.lblFontSize.AutoSize = true;
            this.lblFontSize.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblFontSize.Location = new System.Drawing.Point(20, 35);
            this.lblFontSize.Name = "lblFontSize";
            this.lblFontSize.Size = new System.Drawing.Size(68, 17);
            this.lblFontSize.TabIndex = 0;
            this.lblFontSize.Text = "字体大小：";
            
            this.nudFontSize.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.nudFontSize.Location = new System.Drawing.Point(100, 33);
            this.nudFontSize.Maximum = new decimal(new int[] { 16, 0, 0, 0 });
            this.nudFontSize.Minimum = new decimal(new int[] { 8, 0, 0, 0 });
            this.nudFontSize.Name = "nudFontSize";
            this.nudFontSize.Size = new System.Drawing.Size(60, 23);
            this.nudFontSize.TabIndex = 1;
            this.nudFontSize.Value = new decimal(new int[] { 9, 0, 0, 0 });
            
            this.chkAutoScale.AutoSize = true;
            this.chkAutoScale.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.chkAutoScale.Location = new System.Drawing.Point(20, 70);
            this.chkAutoScale.Name = "chkAutoScale";
            this.chkAutoScale.Size = new System.Drawing.Size(135, 21);
            this.chkAutoScale.TabIndex = 2;
            this.chkAutoScale.Text = "自适应高DPI显示";
            this.chkAutoScale.UseVisualStyleBackColor = true;
            
            // 
            // tabDefaults - 默认值设置
            // 
            this.tabDefaults.Controls.Add(this.grpShippingDefaults);
            this.tabDefaults.Controls.Add(this.grpBillDefaults);
            this.tabDefaults.Location = new System.Drawing.Point(4, 26);
            this.tabDefaults.Name = "tabDefaults";
            this.tabDefaults.Size = new System.Drawing.Size(552, 390);
            this.tabDefaults.TabIndex = 2;
            this.tabDefaults.Text = "📋 默认值";
            this.tabDefaults.UseVisualStyleBackColor = true;
            
            this.grpShippingDefaults.Controls.Add(this.lblDefaultShippingTrack);
            this.grpShippingDefaults.Controls.Add(this.txtDefaultShippingTrack);
            this.grpShippingDefaults.Controls.Add(this.lblDefaultShippingProduct);
            this.grpShippingDefaults.Controls.Add(this.txtDefaultShippingProduct);
            this.grpShippingDefaults.Controls.Add(this.lblDefaultShippingName);
            this.grpShippingDefaults.Controls.Add(this.txtDefaultShippingName);
            this.grpShippingDefaults.Location = new System.Drawing.Point(20, 20);
            this.grpShippingDefaults.Name = "grpShippingDefaults";
            this.grpShippingDefaults.Size = new System.Drawing.Size(250, 140);
            this.grpShippingDefaults.TabIndex = 0;
            this.grpShippingDefaults.TabStop = false;
            this.grpShippingDefaults.Text = "📦 发货明细默认列";
            this.grpShippingDefaults.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            
            this.lblDefaultShippingTrack.AutoSize = true;
            this.lblDefaultShippingTrack.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblDefaultShippingTrack.Location = new System.Drawing.Point(15, 30);
            this.lblDefaultShippingTrack.Name = "lblDefaultShippingTrack";
            this.lblDefaultShippingTrack.Size = new System.Drawing.Size(56, 17);
            this.lblDefaultShippingTrack.TabIndex = 0;
            this.lblDefaultShippingTrack.Text = "运单号：";
            
            this.txtDefaultShippingTrack.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtDefaultShippingTrack.Location = new System.Drawing.Point(80, 27);
            this.txtDefaultShippingTrack.Name = "txtDefaultShippingTrack";
            this.txtDefaultShippingTrack.Size = new System.Drawing.Size(60, 23);
            this.txtDefaultShippingTrack.TabIndex = 1;
            this.txtDefaultShippingTrack.Text = "B";
            
            this.lblDefaultShippingProduct.AutoSize = true;
            this.lblDefaultShippingProduct.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblDefaultShippingProduct.Location = new System.Drawing.Point(15, 65);
            this.lblDefaultShippingProduct.Name = "lblDefaultShippingProduct";
            this.lblDefaultShippingProduct.Size = new System.Drawing.Size(68, 17);
            this.lblDefaultShippingProduct.TabIndex = 2;
            this.lblDefaultShippingProduct.Text = "商品编码：";
            
            this.txtDefaultShippingProduct.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtDefaultShippingProduct.Location = new System.Drawing.Point(80, 62);
            this.txtDefaultShippingProduct.Name = "txtDefaultShippingProduct";
            this.txtDefaultShippingProduct.Size = new System.Drawing.Size(60, 23);
            this.txtDefaultShippingProduct.TabIndex = 3;
            this.txtDefaultShippingProduct.Text = "J";
            
            this.lblDefaultShippingName.AutoSize = true;
            this.lblDefaultShippingName.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblDefaultShippingName.Location = new System.Drawing.Point(15, 100);
            this.lblDefaultShippingName.Name = "lblDefaultShippingName";
            this.lblDefaultShippingName.Size = new System.Drawing.Size(68, 17);
            this.lblDefaultShippingName.TabIndex = 4;
            this.lblDefaultShippingName.Text = "商品名称：";
            
            this.txtDefaultShippingName.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtDefaultShippingName.Location = new System.Drawing.Point(80, 97);
            this.txtDefaultShippingName.Name = "txtDefaultShippingName";
            this.txtDefaultShippingName.Size = new System.Drawing.Size(60, 23);
            this.txtDefaultShippingName.TabIndex = 5;
            this.txtDefaultShippingName.Text = "I";
            
            this.grpBillDefaults.Controls.Add(this.lblDefaultBillTrack);
            this.grpBillDefaults.Controls.Add(this.txtDefaultBillTrack);
            this.grpBillDefaults.Controls.Add(this.lblDefaultBillProduct);
            this.grpBillDefaults.Controls.Add(this.txtDefaultBillProduct);
            this.grpBillDefaults.Controls.Add(this.lblDefaultBillName);
            this.grpBillDefaults.Controls.Add(this.txtDefaultBillName);
            this.grpBillDefaults.Location = new System.Drawing.Point(282, 20);
            this.grpBillDefaults.Name = "grpBillDefaults";
            this.grpBillDefaults.Size = new System.Drawing.Size(250, 140);
            this.grpBillDefaults.TabIndex = 1;
            this.grpBillDefaults.TabStop = false;
            this.grpBillDefaults.Text = "📊 账单明细默认列";
            this.grpBillDefaults.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            
            this.lblDefaultBillTrack.AutoSize = true;
            this.lblDefaultBillTrack.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblDefaultBillTrack.Location = new System.Drawing.Point(15, 30);
            this.lblDefaultBillTrack.Name = "lblDefaultBillTrack";
            this.lblDefaultBillTrack.Size = new System.Drawing.Size(56, 17);
            this.lblDefaultBillTrack.TabIndex = 0;
            this.lblDefaultBillTrack.Text = "运单号：";
            
            this.txtDefaultBillTrack.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtDefaultBillTrack.Location = new System.Drawing.Point(80, 27);
            this.txtDefaultBillTrack.Name = "txtDefaultBillTrack";
            this.txtDefaultBillTrack.Size = new System.Drawing.Size(60, 23);
            this.txtDefaultBillTrack.TabIndex = 1;
            this.txtDefaultBillTrack.Text = "C";
            
            this.lblDefaultBillProduct.AutoSize = true;
            this.lblDefaultBillProduct.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblDefaultBillProduct.Location = new System.Drawing.Point(15, 65);
            this.lblDefaultBillProduct.Name = "lblDefaultBillProduct";
            this.lblDefaultBillProduct.Size = new System.Drawing.Size(68, 17);
            this.lblDefaultBillProduct.TabIndex = 2;
            this.lblDefaultBillProduct.Text = "商品编码：";
            
            this.txtDefaultBillProduct.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtDefaultBillProduct.Location = new System.Drawing.Point(80, 62);
            this.txtDefaultBillProduct.Name = "txtDefaultBillProduct";
            this.txtDefaultBillProduct.Size = new System.Drawing.Size(60, 23);
            this.txtDefaultBillProduct.TabIndex = 3;
            this.txtDefaultBillProduct.Text = "Y";
            
            this.lblDefaultBillName.AutoSize = true;
            this.lblDefaultBillName.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblDefaultBillName.Location = new System.Drawing.Point(15, 100);
            this.lblDefaultBillName.Name = "lblDefaultBillName";
            this.lblDefaultBillName.Size = new System.Drawing.Size(68, 17);
            this.lblDefaultBillName.TabIndex = 4;
            this.lblDefaultBillName.Text = "商品名称：";
            
            this.txtDefaultBillName.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtDefaultBillName.Location = new System.Drawing.Point(80, 97);
            this.txtDefaultBillName.Name = "txtDefaultBillName";
            this.txtDefaultBillName.Size = new System.Drawing.Size(60, 23);
            this.txtDefaultBillName.TabIndex = 5;
            this.txtDefaultBillName.Text = "Z";
            
            // 
            // tabAdvanced - 高级设置
            // 
            this.tabAdvanced.Controls.Add(this.grpAdvanced);
            this.tabAdvanced.Location = new System.Drawing.Point(4, 26);
            this.tabAdvanced.Name = "tabAdvanced";
            this.tabAdvanced.Size = new System.Drawing.Size(552, 390);
            this.tabAdvanced.TabIndex = 3;
            this.tabAdvanced.Text = "⚙️ 高级设置";
            this.tabAdvanced.UseVisualStyleBackColor = true;
            
            this.grpAdvanced.Controls.Add(this.lblLogDirectory);
            this.grpAdvanced.Controls.Add(this.txtLogDirectory);
            this.grpAdvanced.Controls.Add(this.btnBrowseLog);
            this.grpAdvanced.Controls.Add(this.chkAutoSelectSheets);
            this.grpAdvanced.Controls.Add(this.lblProgressUpdate);
            this.grpAdvanced.Controls.Add(this.nudProgressUpdateInterval);
            this.grpAdvanced.Location = new System.Drawing.Point(20, 20);
            this.grpAdvanced.Name = "grpAdvanced";
            this.grpAdvanced.Size = new System.Drawing.Size(512, 160);
            this.grpAdvanced.TabIndex = 0;
            this.grpAdvanced.TabStop = false;
            this.grpAdvanced.Text = "高级选项";
            this.grpAdvanced.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            
            this.lblLogDirectory.AutoSize = true;
            this.lblLogDirectory.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblLogDirectory.Location = new System.Drawing.Point(20, 30);
            this.lblLogDirectory.Name = "lblLogDirectory";
            this.lblLogDirectory.Size = new System.Drawing.Size(68, 17);
            this.lblLogDirectory.TabIndex = 0;
            this.lblLogDirectory.Text = "日志目录：";
            
            this.txtLogDirectory.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtLogDirectory.Location = new System.Drawing.Point(100, 27);
            this.txtLogDirectory.Name = "txtLogDirectory";
            this.txtLogDirectory.Size = new System.Drawing.Size(300, 23);
            this.txtLogDirectory.TabIndex = 1;
            
            this.btnBrowseLog.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.btnBrowseLog.Location = new System.Drawing.Point(410, 26);
            this.btnBrowseLog.Name = "btnBrowseLog";
            this.btnBrowseLog.Size = new System.Drawing.Size(80, 25);
            this.btnBrowseLog.TabIndex = 2;
            this.btnBrowseLog.Text = "浏览...";
            this.btnBrowseLog.UseVisualStyleBackColor = true;
            this.btnBrowseLog.Click += new System.EventHandler(this.btnBrowseLog_Click);
            
            this.chkAutoSelectSheets.AutoSize = true;
            this.chkAutoSelectSheets.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.chkAutoSelectSheets.Location = new System.Drawing.Point(20, 65);
            this.chkAutoSelectSheets.Name = "chkAutoSelectSheets";
            this.chkAutoSelectSheets.Size = new System.Drawing.Size(135, 21);
            this.chkAutoSelectSheets.TabIndex = 3;
            this.chkAutoSelectSheets.Text = "自动选择工作表";
            this.chkAutoSelectSheets.UseVisualStyleBackColor = true;
            
            this.lblProgressUpdate.AutoSize = true;
            this.lblProgressUpdate.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblProgressUpdate.Location = new System.Drawing.Point(20, 100);
            this.lblProgressUpdate.Name = "lblProgressUpdate";
            this.lblProgressUpdate.Size = new System.Drawing.Size(128, 17);
            this.lblProgressUpdate.TabIndex = 4;
            this.lblProgressUpdate.Text = "进度更新间隔(行)：";
            
            this.nudProgressUpdateInterval.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.nudProgressUpdateInterval.Location = new System.Drawing.Point(155, 98);
            this.nudProgressUpdateInterval.Maximum = new decimal(new int[] { 2000, 0, 0, 0 });
            this.nudProgressUpdateInterval.Minimum = new decimal(new int[] { 100, 0, 0, 0 });
            this.nudProgressUpdateInterval.Name = "nudProgressUpdateInterval";
            this.nudProgressUpdateInterval.Size = new System.Drawing.Size(80, 23);
            this.nudProgressUpdateInterval.TabIndex = 5;
            this.nudProgressUpdateInterval.Value = new decimal(new int[] { 500, 0, 0, 0 });
            
            // 
            // 底部按钮
            // 
            this.btnSave.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.btnSave.Location = new System.Drawing.Point(350, 450);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(100, 35);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "💾 保存";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.btnSave.ForeColor = System.Drawing.Color.White;
            this.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            
            this.btnCancel.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.btnCancel.Location = new System.Drawing.Point(470, 450);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(80, 35);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "❌ 取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            
            this.btnReset.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.btnReset.Location = new System.Drawing.Point(20, 450);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(100, 35);
            this.btnReset.TabIndex = 3;
            this.btnReset.Text = "🔄 重置默认";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 500);
            this.Controls.Add(this.tabControl);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnReset);
            this.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "YY运单匹配工具 - 设置";
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(249)))), ((int)(((byte)(250)))));
            this.tabControl.ResumeLayout(false);
            this.tabPerformance.ResumeLayout(false);
            this.tabUI.ResumeLayout(false);
            this.tabDefaults.ResumeLayout(false);
            this.tabAdvanced.ResumeLayout(false);
            this.grpPerformance.ResumeLayout(false);
            this.grpPerformance.PerformLayout();
            this.grpUI.ResumeLayout(false);
            this.grpUI.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudFontSize)).EndInit();
            this.grpShippingDefaults.ResumeLayout(false);
            this.grpShippingDefaults.PerformLayout();
            this.grpBillDefaults.ResumeLayout(false);
            this.grpBillDefaults.PerformLayout();
            this.grpAdvanced.ResumeLayout(false);
            this.grpAdvanced.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudProgressUpdateInterval)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabPerformance;
        private System.Windows.Forms.TabPage tabUI;
        private System.Windows.Forms.TabPage tabDefaults;
        private System.Windows.Forms.TabPage tabAdvanced;
        private System.Windows.Forms.GroupBox grpPerformance;
        private System.Windows.Forms.Label lblPerformanceMode;
        private System.Windows.Forms.ComboBox cmbPerformanceMode;
        private System.Windows.Forms.Label lblPerformanceDesc;
        private System.Windows.Forms.GroupBox grpUI;
        private System.Windows.Forms.Label lblFontSize;
        private System.Windows.Forms.NumericUpDown nudFontSize;
        private System.Windows.Forms.CheckBox chkAutoScale;
        private System.Windows.Forms.GroupBox grpShippingDefaults;
        private System.Windows.Forms.Label lblDefaultShippingTrack;
        private System.Windows.Forms.TextBox txtDefaultShippingTrack;
        private System.Windows.Forms.Label lblDefaultShippingProduct;
        private System.Windows.Forms.TextBox txtDefaultShippingProduct;
        private System.Windows.Forms.Label lblDefaultShippingName;
        private System.Windows.Forms.TextBox txtDefaultShippingName;
        private System.Windows.Forms.GroupBox grpBillDefaults;
        private System.Windows.Forms.Label lblDefaultBillTrack;
        private System.Windows.Forms.TextBox txtDefaultBillTrack;
        private System.Windows.Forms.Label lblDefaultBillProduct;
        private System.Windows.Forms.TextBox txtDefaultBillProduct;
        private System.Windows.Forms.Label lblDefaultBillName;
        private System.Windows.Forms.TextBox txtDefaultBillName;
        private System.Windows.Forms.GroupBox grpAdvanced;
        private System.Windows.Forms.Label lblLogDirectory;
        private System.Windows.Forms.TextBox txtLogDirectory;
        private System.Windows.Forms.Button btnBrowseLog;
        private System.Windows.Forms.CheckBox chkAutoSelectSheets;
        private System.Windows.Forms.Label lblProgressUpdate;
        private System.Windows.Forms.NumericUpDown nudProgressUpdateInterval;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnReset;
    }
} 