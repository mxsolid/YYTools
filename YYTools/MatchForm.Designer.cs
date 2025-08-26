namespace YYTools
{
    partial class MatchForm
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

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.gbShipping = new System.Windows.Forms.GroupBox();
            this.btnSelectNameCol = new System.Windows.Forms.Button();
            this.btnSelectProductCol = new System.Windows.Forms.Button();
            this.btnSelectTrackCol = new System.Windows.Forms.Button();
            this.txtShippingNameColumn = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtShippingProductColumn = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtShippingTrackColumn = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbShippingSheet = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbShippingWorkbook = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.gbBill = new System.Windows.Forms.GroupBox();
            this.btnSelectBillNameCol = new System.Windows.Forms.Button();
            this.btnSelectBillProductCol = new System.Windows.Forms.Button();
            this.btnSelectBillTrackCol = new System.Windows.Forms.Button();
            this.txtBillNameColumn = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtBillProductColumn = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtBillTrackColumn = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.cmbBillSheet = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.cmbBillWorkbook = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.btnSettings = new System.Windows.Forms.Button();
            this.btnViewLogs = new System.Windows.Forms.Button();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.gbShipping.SuspendLayout();
            this.gbBill.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbShipping
            // 
            this.gbShipping.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbShipping.Controls.Add(this.btnSelectNameCol);
            this.gbShipping.Controls.Add(this.btnSelectProductCol);
            this.gbShipping.Controls.Add(this.btnSelectTrackCol);
            this.gbShipping.Controls.Add(this.txtShippingNameColumn);
            this.gbShipping.Controls.Add(this.label5);
            this.gbShipping.Controls.Add(this.txtShippingProductColumn);
            this.gbShipping.Controls.Add(this.label4);
            this.gbShipping.Controls.Add(this.txtShippingTrackColumn);
            this.gbShipping.Controls.Add(this.label3);
            this.gbShipping.Controls.Add(this.cmbShippingSheet);
            this.gbShipping.Controls.Add(this.label2);
            this.gbShipping.Controls.Add(this.cmbShippingWorkbook);
            this.gbShipping.Controls.Add(this.label1);
            this.gbShipping.Location = new System.Drawing.Point(12, 12);
            this.gbShipping.Name = "gbShipping";
            this.gbShipping.Size = new System.Drawing.Size(378, 200);
            this.gbShipping.TabIndex = 0;
            this.gbShipping.TabStop = false;
            this.gbShipping.Text = "🚚 发货明细配置";
            // 
            // btnSelectNameCol
            // 
            this.btnSelectNameCol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectNameCol.Location = new System.Drawing.Point(323, 155);
            this.btnSelectNameCol.Name = "btnSelectNameCol";
            this.btnSelectNameCol.Size = new System.Drawing.Size(40, 25);
            this.btnSelectNameCol.TabIndex = 12;
            this.btnSelectNameCol.Text = "...";
            this.btnSelectNameCol.UseVisualStyleBackColor = true;
            this.btnSelectNameCol.Click += new System.EventHandler(this.btnSelectNameCol_Click);
            // 
            // btnSelectProductCol
            // 
            this.btnSelectProductCol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectProductCol.Location = new System.Drawing.Point(323, 120);
            this.btnSelectProductCol.Name = "btnSelectProductCol";
            this.btnSelectProductCol.Size = new System.Drawing.Size(40, 25);
            this.btnSelectProductCol.TabIndex = 11;
            this.btnSelectProductCol.Text = "...";
            this.btnSelectProductCol.UseVisualStyleBackColor = true;
            this.btnSelectProductCol.Click += new System.EventHandler(this.btnSelectProductCol_Click);
            // 
            // btnSelectTrackCol
            // 
            this.btnSelectTrackCol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectTrackCol.Location = new System.Drawing.Point(323, 85);
            this.btnSelectTrackCol.Name = "btnSelectTrackCol";
            this.btnSelectTrackCol.Size = new System.Drawing.Size(40, 25);
            this.btnSelectTrackCol.TabIndex = 10;
            this.btnSelectTrackCol.Text = "...";
            this.btnSelectTrackCol.UseVisualStyleBackColor = true;
            this.btnSelectTrackCol.Click += new System.EventHandler(this.btnSelectTrackCol_Click);
            // 
            // txtShippingNameColumn
            // 
            this.txtShippingNameColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtShippingNameColumn.Location = new System.Drawing.Point(90, 156);
            this.txtShippingNameColumn.Name = "txtShippingNameColumn";
            this.txtShippingNameColumn.Size = new System.Drawing.Size(227, 23);
            this.txtShippingNameColumn.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(15, 159);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(71, 17);
            this.label5.TabIndex = 8;
            this.label5.Text = "商品名称列:";
            // 
            // txtShippingProductColumn
            // 
            this.txtShippingProductColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtShippingProductColumn.Location = new System.Drawing.Point(90, 121);
            this.txtShippingProductColumn.Name = "txtShippingProductColumn";
            this.txtShippingProductColumn.Size = new System.Drawing.Size(227, 23);
            this.txtShippingProductColumn.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(15, 124);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 17);
            this.label4.TabIndex = 6;
            this.label4.Text = "商品编码列:";
            // 
            // txtShippingTrackColumn
            // 
            this.txtShippingTrackColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtShippingTrackColumn.Location = new System.Drawing.Point(90, 86);
            this.txtShippingTrackColumn.Name = "txtShippingTrackColumn";
            this.txtShippingTrackColumn.Size = new System.Drawing.Size(227, 23);
            this.txtShippingTrackColumn.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 89);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 17);
            this.label3.TabIndex = 4;
            this.label3.Text = "运单号列:";
            // 
            // cmbShippingSheet
            // 
            this.cmbShippingSheet.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbShippingSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbShippingSheet.FormattingEnabled = true;
            this.cmbShippingSheet.Location = new System.Drawing.Point(90, 52);
            this.cmbShippingSheet.Name = "cmbShippingSheet";
            this.cmbShippingSheet.Size = new System.Drawing.Size(273, 25);
            this.cmbShippingSheet.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 17);
            this.label2.TabIndex = 2;
            this.label2.Text = "工作表:";
            // 
            // cmbShippingWorkbook
            // 
            this.cmbShippingWorkbook.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbShippingWorkbook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbShippingWorkbook.FormattingEnabled = true;
            this.cmbShippingWorkbook.Location = new System.Drawing.Point(90, 21);
            this.cmbShippingWorkbook.Name = "cmbShippingWorkbook";
            this.cmbShippingWorkbook.Size = new System.Drawing.Size(273, 25);
            this.cmbShippingWorkbook.TabIndex = 1;
            this.cmbShippingWorkbook.SelectedIndexChanged += new System.EventHandler(this.cmbShippingWorkbook_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "工作簿:";
            // 
            // gbBill
            // 
            this.gbBill.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.gbBill.Controls.Add(this.btnSelectBillNameCol);
            this.gbBill.Controls.Add(this.btnSelectBillProductCol);
            this.gbBill.Controls.Add(this.btnSelectBillTrackCol);
            this.gbBill.Controls.Add(this.txtBillNameColumn);
            this.gbBill.Controls.Add(this.label6);
            this.gbBill.Controls.Add(this.txtBillProductColumn);
            this.gbBill.Controls.Add(this.label7);
            this.gbBill.Controls.Add(this.txtBillTrackColumn);
            this.gbBill.Controls.Add(this.label8);
            this.gbBill.Controls.Add(this.cmbBillSheet);
            this.gbBill.Controls.Add(this.label9);
            this.gbBill.Controls.Add(this.cmbBillWorkbook);
            this.gbBill.Controls.Add(this.label10);
            this.gbBill.Location = new System.Drawing.Point(396, 12);
            this.gbBill.Name = "gbBill";
            this.gbBill.Size = new System.Drawing.Size(378, 200);
            this.gbBill.TabIndex = 1;
            this.gbBill.TabStop = false;
            this.gbBill.Text = "💰 账单明细配置";
            // 
            // btnSelectBillNameCol
            // 
            this.btnSelectBillNameCol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectBillNameCol.Location = new System.Drawing.Point(323, 155);
            this.btnSelectBillNameCol.Name = "btnSelectBillNameCol";
            this.btnSelectBillNameCol.Size = new System.Drawing.Size(40, 25);
            this.btnSelectBillNameCol.TabIndex = 12;
            this.btnSelectBillNameCol.Text = "...";
            this.btnSelectBillNameCol.UseVisualStyleBackColor = true;
            this.btnSelectBillNameCol.Click += new System.EventHandler(this.btnSelectBillNameCol_Click);
            // 
            // btnSelectBillProductCol
            // 
            this.btnSelectBillProductCol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectBillProductCol.Location = new System.Drawing.Point(323, 120);
            this.btnSelectBillProductCol.Name = "btnSelectBillProductCol";
            this.btnSelectBillProductCol.Size = new System.Drawing.Size(40, 25);
            this.btnSelectBillProductCol.TabIndex = 11;
            this.btnSelectBillProductCol.Text = "...";
            this.btnSelectBillProductCol.UseVisualStyleBackColor = true;
            this.btnSelectBillProductCol.Click += new System.EventHandler(this.btnSelectBillProductCol_Click);
            // 
            // btnSelectBillTrackCol
            // 
            this.btnSelectBillTrackCol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectBillTrackCol.Location = new System.Drawing.Point(323, 85);
            this.btnSelectBillTrackCol.Name = "btnSelectBillTrackCol";
            this.btnSelectBillTrackCol.Size = new System.Drawing.Size(40, 25);
            this.btnSelectBillTrackCol.TabIndex = 10;
            this.btnSelectBillTrackCol.Text = "...";
            this.btnSelectBillTrackCol.UseVisualStyleBackColor = true;
            this.btnSelectBillTrackCol.Click += new System.EventHandler(this.btnSelectBillTrackCol_Click);
            // 
            // txtBillNameColumn
            // 
            this.txtBillNameColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtBillNameColumn.Location = new System.Drawing.Point(90, 156);
            this.txtBillNameColumn.Name = "txtBillNameColumn";
            this.txtBillNameColumn.Size = new System.Drawing.Size(227, 23);
            this.txtBillNameColumn.TabIndex = 9;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(15, 159);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(71, 17);
            this.label6.TabIndex = 8;
            this.label6.Text = "商品名称列:";
            // 
            // txtBillProductColumn
            // 
            this.txtBillProductColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtBillProductColumn.Location = new System.Drawing.Point(90, 121);
            this.txtBillProductColumn.Name = "txtBillProductColumn";
            this.txtBillProductColumn.Size = new System.Drawing.Size(227, 23);
            this.txtBillProductColumn.TabIndex = 7;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(15, 124);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(71, 17);
            this.label7.TabIndex = 6;
            this.label7.Text = "商品编码列:";
            // 
            // txtBillTrackColumn
            // 
            this.txtBillTrackColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtBillTrackColumn.Location = new System.Drawing.Point(90, 86);
            this.txtBillTrackColumn.Name = "txtBillTrackColumn";
            this.txtBillTrackColumn.Size = new System.Drawing.Size(227, 23);
            this.txtBillTrackColumn.TabIndex = 5;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(15, 89);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(59, 17);
            this.label8.TabIndex = 4;
            this.label8.Text = "运单号列:";
            // 
            // cmbBillSheet
            // 
            this.cmbBillSheet.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbBillSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBillSheet.FormattingEnabled = true;
            this.cmbBillSheet.Location = new System.Drawing.Point(90, 52);
            this.cmbBillSheet.Name = "cmbBillSheet";
            this.cmbBillSheet.Size = new System.Drawing.Size(273, 25);
            this.cmbBillSheet.TabIndex = 3;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(15, 55);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(47, 17);
            this.label9.TabIndex = 2;
            this.label9.Text = "工作表:";
            // 
            // cmbBillWorkbook
            // 
            this.cmbBillWorkbook.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbBillWorkbook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBillWorkbook.FormattingEnabled = true;
            this.cmbBillWorkbook.Location = new System.Drawing.Point(90, 21);
            this.cmbBillWorkbook.Name = "cmbBillWorkbook";
            this.cmbBillWorkbook.Size = new System.Drawing.Size(273, 25);
            this.cmbBillWorkbook.TabIndex = 1;
            this.cmbBillWorkbook.SelectedIndexChanged += new System.EventHandler(this.cmbBillWorkbook_SelectedIndexChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(15, 24);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(47, 17);
            this.label10.TabIndex = 0;
            this.label10.Text = "工作簿:";
            // 
            // btnStart
            // 
            this.btnStart.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnStart.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.btnStart.FlatAppearance.BorderSize = 0;
            this.btnStart.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnStart.ForeColor = System.Drawing.Color.White;
            this.btnStart.Location = new System.Drawing.Point(578, 259);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(100, 35);
            this.btnStart.TabIndex = 2;
            this.btnStart.Text = "🚀 开始匹配";
            this.btnStart.UseVisualStyleBackColor = false;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Location = new System.Drawing.Point(684, 259);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(90, 35);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "关闭";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(12, 237);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(104, 17);
            this.lblStatus.TabIndex = 4;
            this.lblStatus.Text = "请配置并开始任务";
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(12, 259);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(378, 10);
            this.progressBar.TabIndex = 5;
            this.progressBar.Visible = false;
            // 
            // btnSettings
            // 
            this.btnSettings.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSettings.Location = new System.Drawing.Point(108, 262);
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.Size = new System.Drawing.Size(90, 32);
            this.btnSettings.TabIndex = 6;
            this.btnSettings.Text = "⚙️ 设置";
            this.btnSettings.UseVisualStyleBackColor = true;
            this.btnSettings.Click += new System.EventHandler(this.btnSettings_Click);
            // 
            // btnViewLogs
            // 
            this.btnViewLogs.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnViewLogs.Location = new System.Drawing.Point(12, 262);
            this.btnViewLogs.Name = "btnViewLogs";
            this.btnViewLogs.Size = new System.Drawing.Size(90, 32);
            this.btnViewLogs.TabIndex = 7;
            this.btnViewLogs.Text = "📄 查看日志";
            this.btnViewLogs.UseVisualStyleBackColor = true;
            this.btnViewLogs.Click += new System.EventHandler(this.btnViewLogs_Click);
            // 
            // btnRefresh
            // 
            this.btnRefresh.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefresh.BackColor = System.Drawing.Color.SeaGreen;
            this.btnRefresh.FlatAppearance.BorderSize = 0;
            this.btnRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRefresh.ForeColor = System.Drawing.Color.White;
            this.btnRefresh.Location = new System.Drawing.Point(472, 259);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(100, 35);
            this.btnRefresh.TabIndex = 8;
            this.btnRefresh.Text = "🔄 刷新列表";
            this.btnRefresh.UseVisualStyleBackColor = false;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // MatchForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(786, 306);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.btnViewLogs);
            this.Controls.Add(this.btnSettings);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.gbBill);
            this.Controls.Add(this.gbShipping);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.MinimumSize = new System.Drawing.Size(802, 345);
            this.Name = "MatchForm";
            this.Text = "YY 运单匹配工具 v2.3";
            this.gbShipping.ResumeLayout(false);
            this.gbShipping.PerformLayout();
            this.gbBill.ResumeLayout(false);
            this.gbBill.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox gbShipping;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbShippingWorkbook;
        private System.Windows.Forms.ComboBox cmbShippingSheet;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtShippingTrackColumn;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtShippingNameColumn;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtShippingProductColumn;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnSelectNameCol;
        private System.Windows.Forms.Button btnSelectProductCol;
        private System.Windows.Forms.Button btnSelectTrackCol;
        private System.Windows.Forms.GroupBox gbBill;
        private System.Windows.Forms.Button btnSelectBillNameCol;
        private System.Windows.Forms.Button btnSelectBillProductCol;
        private System.Windows.Forms.Button btnSelectBillTrackCol;
        private System.Windows.Forms.TextBox txtBillNameColumn;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtBillProductColumn;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtBillTrackColumn;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cmbBillSheet;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cmbBillWorkbook;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button btnSettings;
        private System.Windows.Forms.Button btnViewLogs;
        private System.Windows.Forms.Button btnRefresh;
    }
}