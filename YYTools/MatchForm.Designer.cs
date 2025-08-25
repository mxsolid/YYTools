namespace YYTools
{
    partial class MatchForm
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
            this.lblShippingWorkbook = new System.Windows.Forms.Label();
            this.cmbShippingWorkbook = new System.Windows.Forms.ComboBox();
            this.lblShippingSheet = new System.Windows.Forms.Label();
            this.cmbShippingSheet = new System.Windows.Forms.ComboBox();
            this.lblBillWorkbook = new System.Windows.Forms.Label();
            this.cmbBillWorkbook = new System.Windows.Forms.ComboBox();
            this.lblBillSheet = new System.Windows.Forms.Label();
            this.cmbBillSheet = new System.Windows.Forms.ComboBox();
            this.lblShippingTrackColumn = new System.Windows.Forms.Label();
            this.txtShippingTrackColumn = new System.Windows.Forms.TextBox();
            this.lblShippingProductColumn = new System.Windows.Forms.Label();
            this.txtShippingProductColumn = new System.Windows.Forms.TextBox();
            this.lblShippingNameColumn = new System.Windows.Forms.Label();
            this.txtShippingNameColumn = new System.Windows.Forms.TextBox();
            this.lblBillTrackColumn = new System.Windows.Forms.Label();
            this.txtBillTrackColumn = new System.Windows.Forms.TextBox();
            this.lblBillProductColumn = new System.Windows.Forms.Label();
            this.txtBillProductColumn = new System.Windows.Forms.TextBox();
            this.lblBillNameColumn = new System.Windows.Forms.Label();
            this.txtBillNameColumn = new System.Windows.Forms.TextBox();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnViewLogs = new System.Windows.Forms.Button();
            this.btnSettings = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblStatus = new System.Windows.Forms.Label();
            this.grpShipping = new System.Windows.Forms.GroupBox();
            this.grpBill = new System.Windows.Forms.GroupBox();
            this.btnSelectTrackCol = new System.Windows.Forms.Button();
            this.btnSelectProductCol = new System.Windows.Forms.Button();
            this.btnSelectNameCol = new System.Windows.Forms.Button();
            this.btnSelectBillTrackCol = new System.Windows.Forms.Button();
            this.btnSelectBillProductCol = new System.Windows.Forms.Button();
            this.btnSelectBillNameCol = new System.Windows.Forms.Button();
            this.grpShipping.SuspendLayout();
            this.grpBill.SuspendLayout();
            this.SuspendLayout();
            
            // 
            // grpShipping
            // 
            this.grpShipping.Controls.Add(this.lblShippingWorkbook);
            this.grpShipping.Controls.Add(this.cmbShippingWorkbook);
            this.grpShipping.Controls.Add(this.lblShippingSheet);
            this.grpShipping.Controls.Add(this.cmbShippingSheet);
            this.grpShipping.Controls.Add(this.lblShippingTrackColumn);
            this.grpShipping.Controls.Add(this.txtShippingTrackColumn);
            this.grpShipping.Controls.Add(this.btnSelectTrackCol);
            this.grpShipping.Controls.Add(this.lblShippingProductColumn);
            this.grpShipping.Controls.Add(this.txtShippingProductColumn);
            this.grpShipping.Controls.Add(this.btnSelectProductCol);
            this.grpShipping.Controls.Add(this.lblShippingNameColumn);
            this.grpShipping.Controls.Add(this.txtShippingNameColumn);
            this.grpShipping.Controls.Add(this.btnSelectNameCol);
            this.grpShipping.Location = new System.Drawing.Point(20, 20);
            this.grpShipping.Name = "grpShipping";
            this.grpShipping.Size = new System.Drawing.Size(380, 220);
            this.grpShipping.TabIndex = 0;
            this.grpShipping.TabStop = false;
            this.grpShipping.Text = "📦 发货明细设置";
            this.grpShipping.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            
            // 
            // lblShippingWorkbook
            // 
            this.lblShippingWorkbook.AutoSize = true;
            this.lblShippingWorkbook.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblShippingWorkbook.Location = new System.Drawing.Point(15, 25);
            this.lblShippingWorkbook.Name = "lblShippingWorkbook";
            this.lblShippingWorkbook.Size = new System.Drawing.Size(56, 17);
            this.lblShippingWorkbook.TabIndex = 0;
            this.lblShippingWorkbook.Text = "工作簿：";
            
            // 
            // cmbShippingWorkbook
            // 
            this.cmbShippingWorkbook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbShippingWorkbook.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.cmbShippingWorkbook.FormattingEnabled = true;
            this.cmbShippingWorkbook.Location = new System.Drawing.Point(80, 22);
            this.cmbShippingWorkbook.Name = "cmbShippingWorkbook";
            this.cmbShippingWorkbook.Size = new System.Drawing.Size(280, 25);
            this.cmbShippingWorkbook.TabIndex = 1;
            this.cmbShippingWorkbook.SelectedIndexChanged += new System.EventHandler(this.cmbShippingWorkbook_SelectedIndexChanged);
            
            // 
            // lblShippingSheet
            // 
            this.lblShippingSheet.AutoSize = true;
            this.lblShippingSheet.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblShippingSheet.Location = new System.Drawing.Point(15, 55);
            this.lblShippingSheet.Name = "lblShippingSheet";
            this.lblShippingSheet.Size = new System.Drawing.Size(56, 17);
            this.lblShippingSheet.TabIndex = 2;
            this.lblShippingSheet.Text = "工作表：";
            
            // 
            // cmbShippingSheet
            // 
            this.cmbShippingSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbShippingSheet.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.cmbShippingSheet.FormattingEnabled = true;
            this.cmbShippingSheet.Location = new System.Drawing.Point(80, 52);
            this.cmbShippingSheet.Name = "cmbShippingSheet";
            this.cmbShippingSheet.Size = new System.Drawing.Size(280, 25);
            this.cmbShippingSheet.TabIndex = 3;

            // 列配置 - 紧凑布局
            this.lblShippingTrackColumn.AutoSize = true;
            this.lblShippingTrackColumn.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblShippingTrackColumn.Location = new System.Drawing.Point(15, 85);
            this.lblShippingTrackColumn.Name = "lblShippingTrackColumn";
            this.lblShippingTrackColumn.Size = new System.Drawing.Size(68, 17);
            this.lblShippingTrackColumn.TabIndex = 4;
            this.lblShippingTrackColumn.Text = "运单号列：";

            this.txtShippingTrackColumn.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtShippingTrackColumn.Location = new System.Drawing.Point(90, 82);
            this.txtShippingTrackColumn.Name = "txtShippingTrackColumn";
            this.txtShippingTrackColumn.Size = new System.Drawing.Size(200, 23);
            this.txtShippingTrackColumn.TabIndex = 5;
            this.txtShippingTrackColumn.Text = "B";

            this.btnSelectTrackCol.Location = new System.Drawing.Point(300, 81);
            this.btnSelectTrackCol.Name = "btnSelectTrackCol";
            this.btnSelectTrackCol.Size = new System.Drawing.Size(60, 25);
            this.btnSelectTrackCol.TabIndex = 6;
            this.btnSelectTrackCol.Text = "选择";
            this.btnSelectTrackCol.UseVisualStyleBackColor = true;
            this.btnSelectTrackCol.Click += new System.EventHandler(this.btnSelectTrackCol_Click);

            this.lblShippingProductColumn.AutoSize = true;
            this.lblShippingProductColumn.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblShippingProductColumn.Location = new System.Drawing.Point(15, 115);
            this.lblShippingProductColumn.Name = "lblShippingProductColumn";
            this.lblShippingProductColumn.Size = new System.Drawing.Size(68, 17);
            this.lblShippingProductColumn.TabIndex = 7;
            this.lblShippingProductColumn.Text = "商品编码：";

            this.txtShippingProductColumn.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtShippingProductColumn.Location = new System.Drawing.Point(90, 112);
            this.txtShippingProductColumn.Name = "txtShippingProductColumn";
            this.txtShippingProductColumn.Size = new System.Drawing.Size(200, 23);
            this.txtShippingProductColumn.TabIndex = 8;
            this.txtShippingProductColumn.Text = "J";

            this.btnSelectProductCol.Location = new System.Drawing.Point(300, 111);
            this.btnSelectProductCol.Name = "btnSelectProductCol";
            this.btnSelectProductCol.Size = new System.Drawing.Size(60, 25);
            this.btnSelectProductCol.TabIndex = 9;
            this.btnSelectProductCol.Text = "选择";
            this.btnSelectProductCol.UseVisualStyleBackColor = true;
            this.btnSelectProductCol.Click += new System.EventHandler(this.btnSelectProductCol_Click);

            this.lblShippingNameColumn.AutoSize = true;
            this.lblShippingNameColumn.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblShippingNameColumn.Location = new System.Drawing.Point(15, 145);
            this.lblShippingNameColumn.Name = "lblShippingNameColumn";
            this.lblShippingNameColumn.Size = new System.Drawing.Size(68, 17);
            this.lblShippingNameColumn.TabIndex = 10;
            this.lblShippingNameColumn.Text = "商品名称：";

            this.txtShippingNameColumn.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtShippingNameColumn.Location = new System.Drawing.Point(90, 142);
            this.txtShippingNameColumn.Name = "txtShippingNameColumn";
            this.txtShippingNameColumn.Size = new System.Drawing.Size(200, 23);
            this.txtShippingNameColumn.TabIndex = 11;
            this.txtShippingNameColumn.Text = "I";

            this.btnSelectNameCol.Location = new System.Drawing.Point(300, 141);
            this.btnSelectNameCol.Name = "btnSelectNameCol";
            this.btnSelectNameCol.Size = new System.Drawing.Size(60, 25);
            this.btnSelectNameCol.TabIndex = 12;
            this.btnSelectNameCol.Text = "选择";
            this.btnSelectNameCol.UseVisualStyleBackColor = true;
            this.btnSelectNameCol.Click += new System.EventHandler(this.btnSelectNameCol_Click);

            // 
            // grpBill
            // 
            this.grpBill.Controls.Add(this.lblBillWorkbook);
            this.grpBill.Controls.Add(this.cmbBillWorkbook);
            this.grpBill.Controls.Add(this.lblBillSheet);
            this.grpBill.Controls.Add(this.cmbBillSheet);
            this.grpBill.Controls.Add(this.lblBillTrackColumn);
            this.grpBill.Controls.Add(this.txtBillTrackColumn);
            this.grpBill.Controls.Add(this.btnSelectBillTrackCol);
            this.grpBill.Controls.Add(this.lblBillProductColumn);
            this.grpBill.Controls.Add(this.txtBillProductColumn);
            this.grpBill.Controls.Add(this.btnSelectBillProductCol);
            this.grpBill.Controls.Add(this.lblBillNameColumn);
            this.grpBill.Controls.Add(this.txtBillNameColumn);
            this.grpBill.Controls.Add(this.btnSelectBillNameCol);
            this.grpBill.Location = new System.Drawing.Point(420, 20);
            this.grpBill.Name = "grpBill";
            this.grpBill.Size = new System.Drawing.Size(380, 220);
            this.grpBill.TabIndex = 1;
            this.grpBill.TabStop = false;
            this.grpBill.Text = "📊 账单明细设置";
            this.grpBill.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);

            // 账单工作簿和工作表控件 - 紧凑布局
            this.lblBillWorkbook.AutoSize = true;
            this.lblBillWorkbook.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblBillWorkbook.Location = new System.Drawing.Point(15, 25);
            this.lblBillWorkbook.Name = "lblBillWorkbook";
            this.lblBillWorkbook.Size = new System.Drawing.Size(56, 17);
            this.lblBillWorkbook.TabIndex = 0;
            this.lblBillWorkbook.Text = "工作簿：";

            this.cmbBillWorkbook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBillWorkbook.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.cmbBillWorkbook.FormattingEnabled = true;
            this.cmbBillWorkbook.Location = new System.Drawing.Point(80, 22);
            this.cmbBillWorkbook.Name = "cmbBillWorkbook";
            this.cmbBillWorkbook.Size = new System.Drawing.Size(280, 25);
            this.cmbBillWorkbook.TabIndex = 1;
            this.cmbBillWorkbook.SelectedIndexChanged += new System.EventHandler(this.cmbBillWorkbook_SelectedIndexChanged);

            this.lblBillSheet.AutoSize = true;
            this.lblBillSheet.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblBillSheet.Location = new System.Drawing.Point(15, 55);
            this.lblBillSheet.Name = "lblBillSheet";
            this.lblBillSheet.Size = new System.Drawing.Size(56, 17);
            this.lblBillSheet.TabIndex = 2;
            this.lblBillSheet.Text = "工作表：";

            this.cmbBillSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBillSheet.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.cmbBillSheet.FormattingEnabled = true;
            this.cmbBillSheet.Location = new System.Drawing.Point(80, 52);
            this.cmbBillSheet.Name = "cmbBillSheet";
            this.cmbBillSheet.Size = new System.Drawing.Size(280, 25);
            this.cmbBillSheet.TabIndex = 3;

            // 账单列配置 - 紧凑布局
            this.lblBillTrackColumn.AutoSize = true;
            this.lblBillTrackColumn.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblBillTrackColumn.Location = new System.Drawing.Point(15, 85);
            this.lblBillTrackColumn.Name = "lblBillTrackColumn";
            this.lblBillTrackColumn.Size = new System.Drawing.Size(68, 17);
            this.lblBillTrackColumn.TabIndex = 4;
            this.lblBillTrackColumn.Text = "运单号列：";

            this.txtBillTrackColumn.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtBillTrackColumn.Location = new System.Drawing.Point(90, 82);
            this.txtBillTrackColumn.Name = "txtBillTrackColumn";
            this.txtBillTrackColumn.Size = new System.Drawing.Size(200, 23);
            this.txtBillTrackColumn.TabIndex = 5;
            this.txtBillTrackColumn.Text = "C";

            this.btnSelectBillTrackCol.Location = new System.Drawing.Point(300, 81);
            this.btnSelectBillTrackCol.Name = "btnSelectBillTrackCol";
            this.btnSelectBillTrackCol.Size = new System.Drawing.Size(60, 25);
            this.btnSelectBillTrackCol.TabIndex = 6;
            this.btnSelectBillTrackCol.Text = "选择";
            this.btnSelectBillTrackCol.UseVisualStyleBackColor = true;
            this.btnSelectBillTrackCol.Click += new System.EventHandler(this.btnSelectBillTrackCol_Click);

            this.lblBillProductColumn.AutoSize = true;
            this.lblBillProductColumn.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblBillProductColumn.Location = new System.Drawing.Point(15, 115);
            this.lblBillProductColumn.Name = "lblBillProductColumn";
            this.lblBillProductColumn.Size = new System.Drawing.Size(68, 17);
            this.lblBillProductColumn.TabIndex = 7;
            this.lblBillProductColumn.Text = "商品编码：";

            this.txtBillProductColumn.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtBillProductColumn.Location = new System.Drawing.Point(90, 112);
            this.txtBillProductColumn.Name = "txtBillProductColumn";
            this.txtBillProductColumn.Size = new System.Drawing.Size(200, 23);
            this.txtBillProductColumn.TabIndex = 8;
            this.txtBillProductColumn.Text = "Y";

            this.btnSelectBillProductCol.Location = new System.Drawing.Point(300, 111);
            this.btnSelectBillProductCol.Name = "btnSelectBillProductCol";
            this.btnSelectBillProductCol.Size = new System.Drawing.Size(60, 25);
            this.btnSelectBillProductCol.TabIndex = 9;
            this.btnSelectBillProductCol.Text = "选择";
            this.btnSelectBillProductCol.UseVisualStyleBackColor = true;
            this.btnSelectBillProductCol.Click += new System.EventHandler(this.btnSelectBillProductCol_Click);

            this.lblBillNameColumn.AutoSize = true;
            this.lblBillNameColumn.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblBillNameColumn.Location = new System.Drawing.Point(15, 145);
            this.lblBillNameColumn.Name = "lblBillNameColumn";
            this.lblBillNameColumn.Size = new System.Drawing.Size(68, 17);
            this.lblBillNameColumn.TabIndex = 10;
            this.lblBillNameColumn.Text = "商品名称：";

            this.txtBillNameColumn.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.txtBillNameColumn.Location = new System.Drawing.Point(90, 142);
            this.txtBillNameColumn.Name = "txtBillNameColumn";
            this.txtBillNameColumn.Size = new System.Drawing.Size(200, 23);
            this.txtBillNameColumn.TabIndex = 11;
            this.txtBillNameColumn.Text = "Z";

            this.btnSelectBillNameCol.Location = new System.Drawing.Point(300, 141);
            this.btnSelectBillNameCol.Name = "btnSelectBillNameCol";
            this.btnSelectBillNameCol.Size = new System.Drawing.Size(60, 25);
            this.btnSelectBillNameCol.TabIndex = 12;
            this.btnSelectBillNameCol.Text = "选择";
            this.btnSelectBillNameCol.UseVisualStyleBackColor = true;
            this.btnSelectBillNameCol.Click += new System.EventHandler(this.btnSelectBillNameCol_Click);

            // 
            // progressBar - 紧凑布局
            // 
            this.progressBar.Location = new System.Drawing.Point(20, 255);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(780, 20);
            this.progressBar.TabIndex = 2;
            this.progressBar.Visible = false;

            // 
            // lblStatus - 紧凑布局
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.lblStatus.Location = new System.Drawing.Point(20, 285);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(200, 17);
            this.lblStatus.TabIndex = 3;
            this.lblStatus.Text = "准备就绪，请配置工作表和列信息";

            // 
            // 按钮组 - 紧凑布局，间距合理
            // 
            this.btnStart.Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold);
            this.btnStart.Location = new System.Drawing.Point(520, 320);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(130, 40);
            this.btnStart.TabIndex = 4;
            this.btnStart.Text = "🚀 开始匹配";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.btnStart.ForeColor = System.Drawing.Color.White;
            this.btnStart.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);

            this.btnCancel.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.btnCancel.Location = new System.Drawing.Point(670, 320);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(100, 40);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "❌ 取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);

            this.btnViewLogs.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.btnViewLogs.Location = new System.Drawing.Point(20, 320);
            this.btnViewLogs.Name = "btnViewLogs";
            this.btnViewLogs.Size = new System.Drawing.Size(100, 40);
            this.btnViewLogs.TabIndex = 6;
            this.btnViewLogs.Text = "📋 查看日志";
            this.btnViewLogs.UseVisualStyleBackColor = true;
            this.btnViewLogs.Click += new System.EventHandler(this.btnViewLogs_Click);

            this.btnSettings.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.btnSettings.Location = new System.Drawing.Point(120, 320);
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.Size = new System.Drawing.Size(100, 40);
            this.btnSettings.TabIndex = 7;
            this.btnSettings.Text = "⚙️ 设置";
            this.btnSettings.UseVisualStyleBackColor = true;
            this.btnSettings.Click += new System.EventHandler(this.btnSettings_Click);

            // 
            // MatchForm - 紧凑美观
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(820, 380);
            this.Controls.Add(this.grpShipping);
            this.Controls.Add(this.grpBill);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnViewLogs);
            this.Controls.Add(this.btnSettings);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "MatchForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "YY运单匹配工具 v1.5";
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(249)))), ((int)(((byte)(250)))));
            this.grpShipping.ResumeLayout(false);
            this.grpShipping.PerformLayout();
            this.grpBill.ResumeLayout(false);
            this.grpBill.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblShippingWorkbook;
        private System.Windows.Forms.ComboBox cmbShippingWorkbook;
        private System.Windows.Forms.Label lblShippingSheet;
        private System.Windows.Forms.ComboBox cmbShippingSheet;
        private System.Windows.Forms.Label lblBillWorkbook;
        private System.Windows.Forms.ComboBox cmbBillWorkbook;
        private System.Windows.Forms.Label lblBillSheet;
        private System.Windows.Forms.ComboBox cmbBillSheet;
        private System.Windows.Forms.Label lblShippingTrackColumn;
        private System.Windows.Forms.TextBox txtShippingTrackColumn;
        private System.Windows.Forms.Label lblShippingProductColumn;
        private System.Windows.Forms.TextBox txtShippingProductColumn;
        private System.Windows.Forms.Label lblShippingNameColumn;
        private System.Windows.Forms.TextBox txtShippingNameColumn;
        private System.Windows.Forms.Label lblBillTrackColumn;
        private System.Windows.Forms.TextBox txtBillTrackColumn;
        private System.Windows.Forms.Label lblBillProductColumn;
        private System.Windows.Forms.TextBox txtBillProductColumn;
        private System.Windows.Forms.Label lblBillNameColumn;
        private System.Windows.Forms.TextBox txtBillNameColumn;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnViewLogs;
        private System.Windows.Forms.Button btnSettings;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.GroupBox grpShipping;
        private System.Windows.Forms.GroupBox grpBill;
        private System.Windows.Forms.Button btnSelectTrackCol;
        private System.Windows.Forms.Button btnSelectProductCol;
        private System.Windows.Forms.Button btnSelectNameCol;
        private System.Windows.Forms.Button btnSelectBillTrackCol;
        private System.Windows.Forms.Button btnSelectBillProductCol;
        private System.Windows.Forms.Button btnSelectBillNameCol;
    }
} 