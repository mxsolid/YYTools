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
            this.components = new System.ComponentModel.Container();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.refreshListToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.taskOptionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewLogsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.gbShipping = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbShippingNameColumn = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbShippingProductColumn = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbShippingTrackColumn = new System.Windows.Forms.ComboBox();
            this.cmbShippingSheet = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lblShippingWorkbook = new System.Windows.Forms.Label();
            this.cmbShippingWorkbook = new System.Windows.Forms.ComboBox();
            this.gbBill = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cmbBillNameColumn = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.cmbBillProductColumn = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cmbBillTrackColumn = new System.Windows.Forms.ComboBox();
            this.cmbBillSheet = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.cmbBillWorkbook = new System.Windows.Forms.ComboBox();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.panelStatus = new System.Windows.Forms.Panel();
            this.lblStatus = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.gbOptions = new System.Windows.Forms.GroupBox();
            this.cmbSort = new System.Windows.Forms.ComboBox();
            this.chkRemoveDuplicates = new System.Windows.Forms.CheckBox();
            this.txtDelimiter = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.gbWritePreview = new System.Windows.Forms.GroupBox();
            this.txtWritePreview = new System.Windows.Forms.TextBox();
            this.menuStrip1.SuspendLayout();
            this.gbShipping.SuspendLayout();
            this.gbBill.SuspendLayout();
            this.panelStatus.SuspendLayout();
            this.gbOptions.SuspendLayout();
            this.gbWritePreview.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.toolsToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(484, 25);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openFileToolStripMenuItem,
            this.refreshListToolStripMenuItem,
            this.toolStripSeparator1,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(58, 21);
            this.fileToolStripMenuItem.Text = "文件(&F)";
            // 
            // openFileToolStripMenuItem
            // 
            this.openFileToolStripMenuItem.Name = "openFileToolStripMenuItem";
            this.openFileToolStripMenuItem.Size = new System.Drawing.Size(144, 22);
            this.openFileToolStripMenuItem.Text = "打开文件(&O)";
            this.openFileToolStripMenuItem.Click += new System.EventHandler(this.openFileToolStripMenuItem_Click);
            // 
            // refreshListToolStripMenuItem
            // 
            this.refreshListToolStripMenuItem.Name = "refreshListToolStripMenuItem";
            this.refreshListToolStripMenuItem.Size = new System.Drawing.Size(144, 22);
            this.refreshListToolStripMenuItem.Text = "刷新列表(&R)";
            this.refreshListToolStripMenuItem.Click += new System.EventHandler(this.refreshListToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(141, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(144, 22);
            this.exitToolStripMenuItem.Text = "退出(&X)";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // toolsToolStripMenuItem
            // 
            this.toolsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.settingsToolStripMenuItem});
            this.toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            this.toolsToolStripMenuItem.Size = new System.Drawing.Size(59, 21);
            this.toolsToolStripMenuItem.Text = "工具(&T)";
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(116, 22);
            this.settingsToolStripMenuItem.Text = "设置(&S)";
            this.settingsToolStripMenuItem.Click += new System.EventHandler(this.settingsToolStripMenuItem_Click);
            // 
            
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.viewLogsToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(61, 21);
            this.helpToolStripMenuItem.Text = "帮助(&H)";
            // 
            // viewLogsToolStripMenuItem
            // 
            this.viewLogsToolStripMenuItem.Name = "viewLogsToolStripMenuItem";
            this.viewLogsToolStripMenuItem.Size = new System.Drawing.Size(140, 22);
            this.viewLogsToolStripMenuItem.Text = "查看日志(&L)";
            this.viewLogsToolStripMenuItem.Click += new System.EventHandler(this.viewLogsToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(140, 22);
            this.aboutToolStripMenuItem.Text = "关于(&A)";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // gbShipping
            // 
            this.gbShipping.Controls.Add(this.label4);
            this.gbShipping.Controls.Add(this.cmbShippingNameColumn);
            this.gbShipping.Controls.Add(this.label3);
            this.gbShipping.Controls.Add(this.cmbShippingProductColumn);
            this.gbShipping.Controls.Add(this.label2);
            this.gbShipping.Controls.Add(this.cmbShippingTrackColumn);
            this.gbShipping.Controls.Add(this.cmbShippingSheet);
            this.gbShipping.Controls.Add(this.label1);
            this.gbShipping.Controls.Add(this.lblShippingWorkbook);
            this.gbShipping.Controls.Add(this.cmbShippingWorkbook);
            this.gbShipping.Location = new System.Drawing.Point(12, 35);
            this.gbShipping.Name = "gbShipping";
            this.gbShipping.Size = new System.Drawing.Size(460, 185);
            this.gbShipping.TabIndex = 1;
            this.gbShipping.TabStop = false;
            this.gbShipping.Text = "发货明细配置";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(6, 148);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(76, 23);
            this.label4.TabIndex = 9;
            this.label4.Text = "商品名称：";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cmbShippingNameColumn
            // 
            this.cmbShippingNameColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbShippingNameColumn.FormattingEnabled = true;
            this.cmbShippingNameColumn.Location = new System.Drawing.Point(88, 147);
            this.cmbShippingNameColumn.Name = "cmbShippingNameColumn";
            this.cmbShippingNameColumn.Size = new System.Drawing.Size(356, 25);
            this.cmbShippingNameColumn.TabIndex = 4;
            this.cmbShippingNameColumn.SelectedIndexChanged += new System.EventHandler(this.cmbShippingNameColumn_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(6, 117);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(76, 23);
            this.label3.TabIndex = 7;
            this.label3.Text = "商品编码：";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cmbShippingProductColumn
            // 
            this.cmbShippingProductColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbShippingProductColumn.FormattingEnabled = true;
            this.cmbShippingProductColumn.Location = new System.Drawing.Point(88, 116);
            this.cmbShippingProductColumn.Name = "cmbShippingProductColumn";
            this.cmbShippingProductColumn.Size = new System.Drawing.Size(356, 25);
            this.cmbShippingProductColumn.TabIndex = 3;
            this.cmbShippingProductColumn.SelectedIndexChanged += new System.EventHandler(this.cmbShippingProductColumn_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(6, 86);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 23);
            this.label2.TabIndex = 5;
            this.label2.Text = "* 运单号：";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.toolTip1.SetToolTip(this.label2, "此列为匹配核心，将用于在两个表中查找相同的运单");
            // 
            // cmbShippingTrackColumn
            // 
            this.cmbShippingTrackColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbShippingTrackColumn.FormattingEnabled = true;
            this.cmbShippingTrackColumn.Location = new System.Drawing.Point(88, 85);
            this.cmbShippingTrackColumn.Name = "cmbShippingTrackColumn";
            this.cmbShippingTrackColumn.Size = new System.Drawing.Size(356, 25);
            this.cmbShippingTrackColumn.TabIndex = 2;
            this.cmbShippingTrackColumn.SelectedIndexChanged += new System.EventHandler(this.cmbShippingTrackColumn_SelectedIndexChanged);
            // 
            // cmbShippingSheet
            // 
            this.cmbShippingSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbShippingSheet.FormattingEnabled = true;
            this.cmbShippingSheet.Location = new System.Drawing.Point(88, 54);
            this.cmbShippingSheet.Name = "cmbShippingSheet";
            this.cmbShippingSheet.Size = new System.Drawing.Size(356, 25);
            this.cmbShippingSheet.TabIndex = 1;
            this.cmbShippingSheet.SelectedIndexChanged += new System.EventHandler(this.cmbShippingSheet_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(6, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 23);
            this.label1.TabIndex = 2;
            this.label1.Text = "工作表：";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblShippingWorkbook
            // 
            this.lblShippingWorkbook.Location = new System.Drawing.Point(6, 23);
            this.lblShippingWorkbook.Name = "lblShippingWorkbook";
            this.lblShippingWorkbook.Size = new System.Drawing.Size(76, 23);
            this.lblShippingWorkbook.TabIndex = 1;
            this.lblShippingWorkbook.Text = "工作簿：";
            this.lblShippingWorkbook.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cmbShippingWorkbook
            // 
            this.cmbShippingWorkbook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbShippingWorkbook.FormattingEnabled = true;
            this.cmbShippingWorkbook.Location = new System.Drawing.Point(88, 23);
            this.cmbShippingWorkbook.Name = "cmbShippingWorkbook";
            this.cmbShippingWorkbook.Size = new System.Drawing.Size(356, 25);
            this.cmbShippingWorkbook.TabIndex = 0;
            this.cmbShippingWorkbook.SelectedIndexChanged += new System.EventHandler(this.cmbShippingWorkbook_SelectedIndexChanged);
            // 
            // gbBill
            // 
            this.gbBill.Controls.Add(this.label5);
            this.gbBill.Controls.Add(this.cmbBillNameColumn);
            this.gbBill.Controls.Add(this.label6);
            this.gbBill.Controls.Add(this.cmbBillProductColumn);
            this.gbBill.Controls.Add(this.label7);
            this.gbBill.Controls.Add(this.cmbBillTrackColumn);
            this.gbBill.Controls.Add(this.cmbBillSheet);
            this.gbBill.Controls.Add(this.label8);
            this.gbBill.Controls.Add(this.label9);
            this.gbBill.Controls.Add(this.cmbBillWorkbook);
            this.gbBill.Location = new System.Drawing.Point(12, 226);
            this.gbBill.Name = "gbBill";
            this.gbBill.Size = new System.Drawing.Size(460, 185);
            this.gbBill.TabIndex = 2;
            this.gbBill.TabStop = false;
            this.gbBill.Text = "账单明细配置 (数据将写入此表)";
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(6, 148);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(76, 23);
            this.label5.TabIndex = 9;
            this.label5.Text = "商品名称：";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cmbBillNameColumn
            // 
            this.cmbBillNameColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBillNameColumn.FormattingEnabled = true;
            this.cmbBillNameColumn.Location = new System.Drawing.Point(88, 147);
            this.cmbBillNameColumn.Name = "cmbBillNameColumn";
            this.cmbBillNameColumn.Size = new System.Drawing.Size(356, 25);
            this.cmbBillNameColumn.TabIndex = 4;
            this.cmbBillNameColumn.SelectedIndexChanged += new System.EventHandler(this.cmbBillNameColumn_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(6, 117);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(76, 23);
            this.label6.TabIndex = 7;
            this.label6.Text = "商品编码：";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cmbBillProductColumn
            // 
            this.cmbBillProductColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBillProductColumn.FormattingEnabled = true;
            this.cmbBillProductColumn.Location = new System.Drawing.Point(88, 116);
            this.cmbBillProductColumn.Name = "cmbBillProductColumn";
            this.cmbBillProductColumn.Size = new System.Drawing.Size(356, 25);
            this.cmbBillProductColumn.TabIndex = 3;
            this.cmbBillProductColumn.SelectedIndexChanged += new System.EventHandler(this.cmbBillProductColumn_SelectedIndexChanged);
            // 
            // label7
            // 
            this.label7.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(6, 86);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(76, 23);
            this.label7.TabIndex = 5;
            this.label7.Text = "* 运单号：";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.toolTip1.SetToolTip(this.label7, "此列为匹配核心，将用于在两个表中查找相同的运单");
            // 
            // cmbBillTrackColumn
            // 
            this.cmbBillTrackColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBillTrackColumn.FormattingEnabled = true;
            this.cmbBillTrackColumn.Location = new System.Drawing.Point(88, 85);
            this.cmbBillTrackColumn.Name = "cmbBillTrackColumn";
            this.cmbBillTrackColumn.Size = new System.Drawing.Size(356, 25);
            this.cmbBillTrackColumn.TabIndex = 2;
            this.cmbBillTrackColumn.SelectedIndexChanged += new System.EventHandler(this.cmbBillTrackColumn_SelectedIndexChanged);
            // 
            // cmbBillSheet
            // 
            this.cmbBillSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBillSheet.FormattingEnabled = true;
            this.cmbBillSheet.Location = new System.Drawing.Point(88, 54);
            this.cmbBillSheet.Name = "cmbBillSheet";
            this.cmbBillSheet.Size = new System.Drawing.Size(356, 25);
            this.cmbBillSheet.TabIndex = 1;
            this.cmbBillSheet.SelectedIndexChanged += new System.EventHandler(this.cmbBillSheet_SelectedIndexChanged);
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(6, 54);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(76, 23);
            this.label8.TabIndex = 2;
            this.label8.Text = "工作表：";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label9
            // 
            this.label9.Location = new System.Drawing.Point(6, 23);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(76, 23);
            this.label9.TabIndex = 1;
            this.label9.Text = "工作簿：";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cmbBillWorkbook
            // 
            this.cmbBillWorkbook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBillWorkbook.FormattingEnabled = true;
            this.cmbBillWorkbook.Location = new System.Drawing.Point(88, 23);
            this.cmbBillWorkbook.Name = "cmbBillWorkbook";
            this.cmbBillWorkbook.Size = new System.Drawing.Size(356, 25);
            this.cmbBillWorkbook.TabIndex = 0;
            this.cmbBillWorkbook.SelectedIndexChanged += new System.EventHandler(this.cmbBillWorkbook_SelectedIndexChanged);
            // 
            // btnStart
            // 
            this.btnStart.Font = new System.Drawing.Font("微软雅黑", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnStart.ForeColor = System.Drawing.Color.ForestGreen;
            this.btnStart.Location = new System.Drawing.Point(259, 582);
            this.btnStart.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(100, 30);
            this.btnStart.TabIndex = 5;
            this.btnStart.Text = "▶️ 开始任务";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(372, 582);
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(100, 30);
            this.btnClose.TabIndex = 6;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // panelStatus
            // 
            this.panelStatus.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelStatus.Location = new System.Drawing.Point(0, 604);
            this.panelStatus.Name = "panelStatus";
            this.panelStatus.Size = new System.Drawing.Size(484, 40);
            this.panelStatus.TabIndex = 100;
            this.panelStatus.BackColor = System.Drawing.Color.FromArgb(248,248,248);
            this.panelStatus.Padding = new System.Windows.Forms.Padding(10, 6, 10, 6);
            this.panelStatus.Controls.Add(this.progressBar);
            this.panelStatus.Controls.Add(this.lblStatus);
            // 
            // lblStatus
            // 
            this.lblStatus.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblStatus.AutoSize = false;
            this.lblStatus.Height = 20;
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(464, 20);
            this.lblStatus.TabIndex = 8;
            this.lblStatus.Text = "欢迎使用YY匹配工具";
            // 
            // progressBar
            // 
            this.progressBar.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar.Height = 8;
            this.progressBar.Name = "progressBar";
            this.progressBar.TabIndex = 7;
            // 
            // gbOptions
            // 
            this.gbOptions.Controls.Add(this.cmbSort);
            // 隐藏“去重”按钮，避免干扰输入预览性能
            this.chkRemoveDuplicates.Visible = false;
            this.gbOptions.Controls.Add(this.chkRemoveDuplicates);
            this.gbOptions.Controls.Add(this.txtDelimiter);
            this.gbOptions.Controls.Add(this.label13);
            this.gbOptions.Location = new System.Drawing.Point(12, 417);
            this.gbOptions.Name = "gbOptions";
            this.gbOptions.Size = new System.Drawing.Size(460, 65);
            this.gbOptions.TabIndex = 3;
            this.gbOptions.TabStop = false;
            this.gbOptions.Text = "任务配置";
            // 
            // cmbSort
            // 
            this.cmbSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSort.FormattingEnabled = true;
            this.cmbSort.Items.AddRange(new object[] {
            "默认排序",
            "升序",
            "降序"});
            this.cmbSort.Location = new System.Drawing.Point(324, 25);
            this.cmbSort.Name = "cmbSort";
            this.cmbSort.Size = new System.Drawing.Size(120, 25);
            this.cmbSort.TabIndex = 2;
            this.toolTip1.SetToolTip(this.cmbSort, "对拼接的多个商品编码或名称进行排序");
            // 
            // chkRemoveDuplicates
            // 
            this.chkRemoveDuplicates.AutoSize = true;
            this.chkRemoveDuplicates.Location = new System.Drawing.Point(179, 27);
            this.chkRemoveDuplicates.Name = "chkRemoveDuplicates";
            this.chkRemoveDuplicates.Size = new System.Drawing.Size(123, 21);
            this.chkRemoveDuplicates.TabIndex = 1;
            this.chkRemoveDuplicates.Text = "拼接时自动去重";
            this.toolTip1.SetToolTip(this.chkRemoveDuplicates, "如果匹配到多个相同的商品，只保留一个");
            this.chkRemoveDuplicates.UseVisualStyleBackColor = true;
            // 
            // txtDelimiter
            // 
            this.txtDelimiter.Location = new System.Drawing.Point(88, 25);
            this.txtDelimiter.Name = "txtDelimiter";
            this.txtDelimiter.Size = new System.Drawing.Size(65, 23);
            this.txtDelimiter.TabIndex = 0;
            this.toolTip1.SetToolTip(this.txtDelimiter, "当一个运单匹配到多个商品时，用此符号连接");
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(12, 28);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(80, 17);
            this.label13.TabIndex = 0;
            this.label13.Text = "拼接分隔符：";
            // 
            // gbWritePreview
            // 
            this.gbWritePreview.Controls.Add(this.txtWritePreview);
            this.gbWritePreview.Location = new System.Drawing.Point(12, 488);
            this.gbWritePreview.Name = "gbWritePreview";
            this.gbWritePreview.Size = new System.Drawing.Size(460, 80);
            this.gbWritePreview.TabIndex = 4;
            this.gbWritePreview.TabStop = false;
            this.gbWritePreview.Text = "写入效果预览";
            // 
            // txtWritePreview
            // 
            this.txtWritePreview.BackColor = System.Drawing.SystemColors.Info;
            this.txtWritePreview.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtWritePreview.Location = new System.Drawing.Point(15, 22);
            this.txtWritePreview.Multiline = true;
            this.txtWritePreview.Name = "txtWritePreview";
            this.txtWritePreview.ReadOnly = true;
            this.txtWritePreview.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtWritePreview.Size = new System.Drawing.Size(429, 45);
            this.txtWritePreview.TabIndex = 0;
            // 
            // MatchForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 621);
            this.Controls.Add(this.gbWritePreview);
            this.Controls.Add(this.gbOptions);
            this.Controls.Add(this.panelStatus);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.gbBill);
            this.Controls.Add(this.gbShipping);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = true;
            this.MinimumSize = new System.Drawing.Size(500, 650);
            this.Name = "MatchForm";
            this.Text = "YY 运单匹配工具 v3.1 (性能优化版)";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.gbShipping.ResumeLayout(false);
            this.gbShipping.PerformLayout();
            this.gbBill.ResumeLayout(false);
            this.gbBill.PerformLayout();
            this.gbOptions.ResumeLayout(false);
            this.gbOptions.PerformLayout();
            this.panelStatus.ResumeLayout(false);
            this.gbWritePreview.ResumeLayout(false);
            this.gbWritePreview.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem taskOptionsToolStripMenuItem;
        private System.Windows.Forms.GroupBox gbShipping;
        private System.Windows.Forms.ComboBox cmbShippingWorkbook;
        private System.Windows.Forms.Label lblShippingWorkbook;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbShippingSheet;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbShippingTrackColumn;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbShippingProductColumn;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbShippingNameColumn;
        private System.Windows.Forms.GroupBox gbBill;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cmbBillNameColumn;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cmbBillProductColumn;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmbBillTrackColumn;
        private System.Windows.Forms.ComboBox cmbBillSheet;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cmbBillWorkbook;
        private System.Windows.Forms.ToolStripMenuItem openFileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem refreshListToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
        
        private System.Windows.Forms.ToolStripMenuItem viewLogsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.GroupBox gbOptions;
        private System.Windows.Forms.ComboBox cmbSort;
        private System.Windows.Forms.CheckBox chkRemoveDuplicates;
        private System.Windows.Forms.TextBox txtDelimiter;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.GroupBox gbWritePreview;
        private System.Windows.Forms.TextBox txtWritePreview;
        private System.Windows.Forms.Panel panelStatus;
    }
}