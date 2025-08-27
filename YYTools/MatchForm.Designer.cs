// --- 文件 6: MatchForm.Designer.cs ---
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
            this.lblShippingInfo = new System.Windows.Forms.Label();
            this.cmbShippingNameColumn = new System.Windows.Forms.ComboBox();
            this.cmbShippingProductColumn = new System.Windows.Forms.ComboBox();
            this.cmbShippingTrackColumn = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbShippingSheet = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbShippingWorkbook = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.gbBill = new System.Windows.Forms.GroupBox();
            this.lblBillInfo = new System.Windows.Forms.Label();
            this.cmbBillNameColumn = new System.Windows.Forms.ComboBox();
            this.cmbBillProductColumn = new System.Windows.Forms.ComboBox();
            this.cmbBillTrackColumn = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.cmbBillSheet = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.cmbBillWorkbook = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.refreshListToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewLogsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.bottomPanel = new System.Windows.Forms.Panel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.tabControlMain = new System.Windows.Forms.TabControl();
            this.tabPageMatcher = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.chkRemoveDuplicates = new System.Windows.Forms.CheckBox();
            this.label12 = new System.Windows.Forms.Label();
            this.txtDelimiter = new System.Windows.Forms.TextBox();
            this.tabPageFuture = new System.Windows.Forms.TabPage();
            this.gbShipping.SuspendLayout();
            this.gbBill.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.bottomPanel.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.tabControlMain.SuspendLayout();
            this.tabPageMatcher.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbShipping
            // 
            this.gbShipping.Controls.Add(this.lblShippingInfo);
            this.gbShipping.Controls.Add(this.cmbShippingNameColumn);
            this.gbShipping.Controls.Add(this.cmbShippingProductColumn);
            this.gbShipping.Controls.Add(this.cmbShippingTrackColumn);
            this.gbShipping.Controls.Add(this.label5);
            this.gbShipping.Controls.Add(this.label4);
            this.gbShipping.Controls.Add(this.label3);
            this.gbShipping.Controls.Add(this.cmbShippingSheet);
            this.gbShipping.Controls.Add(this.label2);
            this.gbShipping.Controls.Add(this.cmbShippingWorkbook);
            this.gbShipping.Controls.Add(this.label1);
            this.gbShipping.Location = new System.Drawing.Point(6, 6);
            this.gbShipping.Name = "gbShipping";
            this.gbShipping.Size = new System.Drawing.Size(370, 205);
            this.gbShipping.TabIndex = 0;
            this.gbShipping.TabStop = false;
            this.gbShipping.Text = "🚚 发货明细配置";
            // 
            // lblShippingInfo
            // 
            this.lblShippingInfo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.lblShippingInfo.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.lblShippingInfo.Location = new System.Drawing.Point(15, 185);
            this.lblShippingInfo.Name = "lblShippingInfo";
            this.lblShippingInfo.Size = new System.Drawing.Size(340, 17);
            this.lblShippingInfo.TabIndex = 17;
            // 
            // cmbShippingNameColumn
            // 
            this.cmbShippingNameColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbShippingNameColumn.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cmbShippingNameColumn.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbShippingNameColumn.FormattingEnabled = true;
            this.cmbShippingNameColumn.Location = new System.Drawing.Point(90, 150);
            this.cmbShippingNameColumn.Name = "cmbShippingNameColumn";
            this.cmbShippingNameColumn.Size = new System.Drawing.Size(265, 25);
            this.cmbShippingNameColumn.TabIndex = 16;
            // 
            // cmbShippingProductColumn
            // 
            this.cmbShippingProductColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbShippingProductColumn.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cmbShippingProductColumn.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbShippingProductColumn.FormattingEnabled = true;
            this.cmbShippingProductColumn.Location = new System.Drawing.Point(90, 117);
            this.cmbShippingProductColumn.Name = "cmbShippingProductColumn";
            this.cmbShippingProductColumn.Size = new System.Drawing.Size(265, 25);
            this.cmbShippingProductColumn.TabIndex = 15;
            // 
            // cmbShippingTrackColumn
            // 
            this.cmbShippingTrackColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbShippingTrackColumn.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cmbShippingTrackColumn.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbShippingTrackColumn.FormattingEnabled = true;
            this.cmbShippingTrackColumn.Location = new System.Drawing.Point(90, 84);
            this.cmbShippingTrackColumn.Name = "cmbShippingTrackColumn";
            this.cmbShippingTrackColumn.Size = new System.Drawing.Size(265, 25);
            this.cmbShippingTrackColumn.TabIndex = 14;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(15, 153);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(71, 17);
            this.label5.TabIndex = 8;
            this.label5.Text = "商品名称列:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(15, 120);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 17);
            this.label4.TabIndex = 6;
            this.label4.Text = "商品编码列:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 87);
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
            this.cmbShippingSheet.Location = new System.Drawing.Point(90, 51);
            this.cmbShippingSheet.Name = "cmbShippingSheet";
            this.cmbShippingSheet.Size = new System.Drawing.Size(265, 25);
            this.cmbShippingSheet.TabIndex = 3;
            this.cmbShippingSheet.SelectedIndexChanged += new System.EventHandler(this.cmbShippingSheet_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 54);
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
            this.cmbShippingWorkbook.Location = new System.Drawing.Point(90, 18);
            this.cmbShippingWorkbook.Name = "cmbShippingWorkbook";
            this.cmbShippingWorkbook.Size = new System.Drawing.Size(265, 25);
            this.cmbShippingWorkbook.TabIndex = 1;
            this.cmbShippingWorkbook.SelectedIndexChanged += new System.EventHandler(this.cmbShippingWorkbook_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "工作簿:";
            // 
            // gbBill
            // 
            this.gbBill.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.gbBill.Controls.Add(this.lblBillInfo);
            this.gbBill.Controls.Add(this.cmbBillNameColumn);
            this.gbBill.Controls.Add(this.cmbBillProductColumn);
            this.gbBill.Controls.Add(this.cmbBillTrackColumn);
            this.gbBill.Controls.Add(this.label6);
            this.gbBill.Controls.Add(this.label7);
            this.gbBill.Controls.Add(this.label8);
            this.gbBill.Controls.Add(this.cmbBillSheet);
            this.gbBill.Controls.Add(this.label9);
            this.gbBill.Controls.Add(this.cmbBillWorkbook);
            this.gbBill.Controls.Add(this.label10);
            this.gbBill.Location = new System.Drawing.Point(382, 6);
            this.gbBill.Name = "gbBill";
            this.gbBill.Size = new System.Drawing.Size(370, 205);
            this.gbBill.TabIndex = 1;
            this.gbBill.TabStop = false;
            this.gbBill.Text = "💰 账单明细配置";
            // 
            // lblBillInfo
            // 
            this.lblBillInfo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.lblBillInfo.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.lblBillInfo.Location = new System.Drawing.Point(15, 185);
            this.lblBillInfo.Name = "lblBillInfo";
            this.lblBillInfo.Size = new System.Drawing.Size(340, 17);
            this.lblBillInfo.TabIndex = 18;
            // 
            // cmbBillNameColumn
            // 
            this.cmbBillNameColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbBillNameColumn.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cmbBillNameColumn.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbBillNameColumn.FormattingEnabled = true;
            this.cmbBillNameColumn.Location = new System.Drawing.Point(90, 150);
            this.cmbBillNameColumn.Name = "cmbBillNameColumn";
            this.cmbBillNameColumn.Size = new System.Drawing.Size(265, 25);
            this.cmbBillNameColumn.TabIndex = 17;
            // 
            // cmbBillProductColumn
            // 
            this.cmbBillProductColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbBillProductColumn.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cmbBillProductColumn.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbBillProductColumn.FormattingEnabled = true;
            this.cmbBillProductColumn.Location = new System.Drawing.Point(90, 117);
            this.cmbBillProductColumn.Name = "cmbBillProductColumn";
            this.cmbBillProductColumn.Size = new System.Drawing.Size(265, 25);
            this.cmbBillProductColumn.TabIndex = 16;
            // 
            // cmbBillTrackColumn
            // 
            this.cmbBillTrackColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbBillTrackColumn.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cmbBillTrackColumn.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbBillTrackColumn.FormattingEnabled = true;
            this.cmbBillTrackColumn.Location = new System.Drawing.Point(90, 84);
            this.cmbBillTrackColumn.Name = "cmbBillTrackColumn";
            this.cmbBillTrackColumn.Size = new System.Drawing.Size(265, 25);
            this.cmbBillTrackColumn.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(15, 153);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(71, 17);
            this.label6.TabIndex = 8;
            this.label6.Text = "商品名称列:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(15, 120);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(71, 17);
            this.label7.TabIndex = 6;
            this.label7.Text = "商品编码列:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(15, 87);
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
            this.cmbBillSheet.Location = new System.Drawing.Point(90, 51);
            this.cmbBillSheet.Name = "cmbBillSheet";
            this.cmbBillSheet.Size = new System.Drawing.Size(265, 25);
            this.cmbBillSheet.TabIndex = 3;
            this.cmbBillSheet.SelectedIndexChanged += new System.EventHandler(this.cmbBillSheet_SelectedIndexChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(15, 54);
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
            this.cmbBillWorkbook.Location = new System.Drawing.Point(90, 18);
            this.cmbBillWorkbook.Name = "cmbBillWorkbook";
            this.cmbBillWorkbook.Size = new System.Drawing.Size(265, 25);
            this.cmbBillWorkbook.TabIndex = 1;
            this.cmbBillWorkbook.SelectedIndexChanged += new System.EventHandler(this.cmbBillWorkbook_SelectedIndexChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(15, 21);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(47, 17);
            this.label10.TabIndex = 0;
            this.label10.Text = "工作簿:";
            // 
            // lblStatus
            // 
            this.lblStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.lblStatus.Location = new System.Drawing.Point(8, 7);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(760, 17);
            this.lblStatus.TabIndex = 4;
            this.lblStatus.Text = "请配置并开始任务";
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(11, 29);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(757, 10);
            this.progressBar.TabIndex = 5;
            this.progressBar.Visible = false;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.toolsToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(784, 25);
            this.menuStrip1.TabIndex = 9;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openFileToolStripMenuItem,
            this.refreshListToolStripMenuItem,
            this.toolStripMenuItem1,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(58, 21);
            this.fileToolStripMenuItem.Text = "文件(&F)";
            // 
            // openFileToolStripMenuItem
            // 
            this.openFileToolStripMenuItem.Name = "openFileToolStripMenuItem";
            this.openFileToolStripMenuItem.Size = new System.Drawing.Size(144, 22);
            this.openFileToolStripMenuItem.Text = "打开文件(&O)...";
            this.openFileToolStripMenuItem.Click += new System.EventHandler(this.openFileToolStripMenuItem_Click);
            // 
            // refreshListToolStripMenuItem
            // 
            this.refreshListToolStripMenuItem.Name = "refreshListToolStripMenuItem";
            this.refreshListToolStripMenuItem.Size = new System.Drawing.Size(144, 22);
            this.refreshListToolStripMenuItem.Text = "刷新列表(&R)";
            this.refreshListToolStripMenuItem.Click += new System.EventHandler(this.refreshListToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(141, 6);
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
            this.settingsToolStripMenuItem,
            this.viewLogsToolStripMenuItem});
            this.toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            this.toolsToolStripMenuItem.Size = new System.Drawing.Size(59, 21);
            this.toolsToolStripMenuItem.Text = "工具(&T)";
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(140, 22);
            this.settingsToolStripMenuItem.Text = "设置(&S)...";
            this.settingsToolStripMenuItem.Click += new System.EventHandler(this.settingsToolStripMenuItem_Click);
            // 
            // viewLogsToolStripMenuItem
            // 
            this.viewLogsToolStripMenuItem.Name = "viewLogsToolStripMenuItem";
            this.viewLogsToolStripMenuItem.Size = new System.Drawing.Size(140, 22);
            this.viewLogsToolStripMenuItem.Text = "查看日志(&L)";
            this.viewLogsToolStripMenuItem.Click += new System.EventHandler(this.viewLogsToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(61, 21);
            this.helpToolStripMenuItem.Text = "帮助(&H)";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.aboutToolStripMenuItem.Text = "关于(&A)...";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // bottomPanel
            // 
            this.bottomPanel.Controls.Add(this.flowLayoutPanel1);
            this.bottomPanel.Controls.Add(this.progressBar);
            this.bottomPanel.Controls.Add(this.lblStatus);
            this.bottomPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.bottomPanel.Location = new System.Drawing.Point(0, 292);
            this.bottomPanel.Name = "bottomPanel";
            this.bottomPanel.Padding = new System.Windows.Forms.Padding(8, 0, 8, 8);
            this.bottomPanel.Size = new System.Drawing.Size(784, 89);
            this.bottomPanel.TabIndex = 10;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.btnClose);
            this.flowLayoutPanel1.Controls.Add(this.btnStart);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.flowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(8, 43);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Padding = new System.Windows.Forms.Padding(0, 5, 0, 0);
            this.flowLayoutPanel1.Size = new System.Drawing.Size(768, 38);
            this.flowLayoutPanel1.TabIndex = 6;
            // 
            // btnClose
            // 
            this.btnClose.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnClose.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnClose.Location = new System.Drawing.Point(645, 8);
            this.btnClose.Margin = new System.Windows.Forms.Padding(3, 3, 10, 3);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(113, 27);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnStart
            // 
            this.btnStart.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.btnStart.FlatAppearance.BorderSize = 0;
            this.btnStart.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnStart.ForeColor = System.Drawing.Color.White;
            this.btnStart.Location = new System.Drawing.Point(526, 8);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(113, 27);
            this.btnStart.TabIndex = 3;
            this.btnStart.Text = "🚀 开始匹配";
            this.btnStart.UseVisualStyleBackColor = false;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // tabControlMain
            // 
            this.tabControlMain.Controls.Add(this.tabPageMatcher);
            this.tabControlMain.Controls.Add(this.tabPageFuture);
            this.tabControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlMain.Location = new System.Drawing.Point(0, 25);
            this.tabControlMain.Name = "tabControlMain";
            this.tabControlMain.SelectedIndex = 0;
            this.tabControlMain.Size = new System.Drawing.Size(784, 267);
            this.tabControlMain.TabIndex = 11;
            // 
            // tabPageMatcher
            // 
            this.tabPageMatcher.BackColor = System.Drawing.SystemColors.Control;
            this.tabPageMatcher.Controls.Add(this.groupBox3);
            this.tabPageMatcher.Controls.Add(this.gbShipping);
            this.tabPageMatcher.Controls.Add(this.gbBill);
            this.tabPageMatcher.Location = new System.Drawing.Point(4, 26);
            this.tabPageMatcher.Name = "tabPageMatcher";
            this.tabPageMatcher.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageMatcher.Size = new System.Drawing.Size(776, 237);
            this.tabPageMatcher.TabIndex = 0;
            this.tabPageMatcher.Text = "运单匹配";
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox3.Controls.Add(this.chkRemoveDuplicates);
            this.groupBox3.Controls.Add(this.label12);
            this.groupBox3.Controls.Add(this.txtDelimiter);
            this.groupBox3.Location = new System.Drawing.Point(6, 217);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(746, 50);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "匹配选项";
            // 
            // chkRemoveDuplicates
            // 
            this.chkRemoveDuplicates.AutoSize = true;
            this.chkRemoveDuplicates.Location = new System.Drawing.Point(176, 22);
            this.chkRemoveDuplicates.Name = "chkRemoveDuplicates";
            this.chkRemoveDuplicates.Size = new System.Drawing.Size(87, 21);
            this.chkRemoveDuplicates.TabIndex = 2;
            this.chkRemoveDuplicates.Text = "拼接时去重";
            this.chkRemoveDuplicates.UseVisualStyleBackColor = true;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(15, 24);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(47, 17);
            this.label12.TabIndex = 1;
            this.label12.Text = "拼接符:";
            // 
            // txtDelimiter
            // 
            this.txtDelimiter.Location = new System.Drawing.Point(68, 21);
            this.txtDelimiter.Name = "txtDelimiter";
            this.txtDelimiter.Size = new System.Drawing.Size(80, 23);
            this.txtDelimiter.TabIndex = 0;
            // 
            // tabPageFuture
            // 
            this.tabPageFuture.BackColor = System.Drawing.SystemColors.Control;
            this.tabPageFuture.Location = new System.Drawing.Point(4, 26);
            this.tabPageFuture.Name = "tabPageFuture";
            this.tabPageFuture.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageFuture.Size = new System.Drawing.Size(776, 277);
            this.tabPageFuture.TabIndex = 1;
            this.tabPageFuture.Text = "（未来工具）";
            // 
            // MatchForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(784, 381);
            this.Controls.Add(this.tabControlMain);
            this.Controls.Add(this.bottomPanel);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(800, 420);
            this.Name = "MatchForm";
            this.Text = "YY 运单匹配工具 v2.5";
            this.gbShipping.ResumeLayout(false);
            this.gbShipping.PerformLayout();
            this.gbBill.ResumeLayout(false);
            this.gbBill.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.bottomPanel.ResumeLayout(false);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.tabControlMain.ResumeLayout(false);
            this.tabPageMatcher.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox gbShipping;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbShippingWorkbook;
        private System.Windows.Forms.ComboBox cmbShippingSheet;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox gbBill;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cmbBillSheet;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cmbBillWorkbook;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openFileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem refreshListToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewLogsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.Panel bottomPanel;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.ComboBox cmbShippingTrackColumn;
        private System.Windows.Forms.ComboBox cmbShippingNameColumn;
        private System.Windows.Forms.ComboBox cmbShippingProductColumn;
        private System.Windows.Forms.ComboBox cmbBillNameColumn;
        private System.Windows.Forms.ComboBox cmbBillProductColumn;
        private System.Windows.Forms.ComboBox cmbBillTrackColumn;
        private System.Windows.Forms.TabControl tabControlMain;
        private System.Windows.Forms.TabPage tabPageMatcher;
        private System.Windows.Forms.TabPage tabPageFuture;
        private System.Windows.Forms.Label lblShippingInfo;
        private System.Windows.Forms.Label lblBillInfo;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckBox chkRemoveDuplicates;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtDelimiter;
    }
}