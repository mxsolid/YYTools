// --- MatchForm.Designer.cs ---
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
            this.lblBillInfo = new System.Windows.Forms.Label();
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
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.gbShipping.SuspendLayout();
            this.gbBill.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.bottomPanel.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbShipping
            // 
            this.gbShipping.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbShipping.Controls.Add(this.lblShippingInfo);
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
            this.gbShipping.Location = new System.Drawing.Point(12, 38);
            this.gbShipping.Name = "gbShipping";
            this.gbShipping.Size = new System.Drawing.Size(378, 230);
            this.gbShipping.TabIndex = 0;
            this.gbShipping.TabStop = false;
            this.gbShipping.Text = "🚚 发货明细配置";
            // 
            // lblShippingInfo
            // 
            this.lblShippingInfo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblShippingInfo.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.lblShippingInfo.Location = new System.Drawing.Point(90, 131);
            this.lblShippingInfo.Name = "lblShippingInfo";
            this.lblShippingInfo.Size = new System.Drawing.Size(273, 17);
            this.lblShippingInfo.TabIndex = 13;
            this.lblShippingInfo.Text = "总行数: - | 文件大小: -";
            // 
            // btnSelectNameCol
            // 
            this.btnSelectNameCol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectNameCol.Location = new System.Drawing.Point(323, 186);
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
            this.btnSelectProductCol.Location = new System.Drawing.Point(323, 151);
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
            this.txtShippingNameColumn.Location = new System.Drawing.Point(90, 187);
            this.txtShippingNameColumn.Name = "txtShippingNameColumn";
            this.txtShippingNameColumn.Size = new System.Drawing.Size(227, 23);
            this.txtShippingNameColumn.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(15, 190);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(71, 17);
            this.label5.TabIndex = 8;
            this.label5.Text = "商品名称列:";
            // 
            // txtShippingProductColumn
            // 
            this.txtShippingProductColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtShippingProductColumn.Location = new System.Drawing.Point(90, 152);
            this.txtShippingProductColumn.Name = "txtShippingProductColumn";
            this.txtShippingProductColumn.Size = new System.Drawing.Size(227, 23);
            this.txtShippingProductColumn.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(15, 155);
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
            this.gbBill.Controls.Add(this.lblBillInfo);
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
            this.gbBill.Location = new System.Drawing.Point(396, 38);
            this.gbBill.Name = "gbBill";
            this.gbBill.Size = new System.Drawing.Size(378, 230);
            this.gbBill.TabIndex = 1;
            this.gbBill.TabStop = false;
            this.gbBill.Text = "💰 账单明细配置";
            // 
            // lblBillInfo
            // 
            this.lblBillInfo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblBillInfo.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.lblBillInfo.Location = new System.Drawing.Point(90, 131);
            this.lblBillInfo.Name = "lblBillInfo";
            this.lblBillInfo.Size = new System.Drawing.Size(273, 17);
            this.lblBillInfo.TabIndex = 14;
            this.lblBillInfo.Text = "总行数: - | 文件大小: -";
            // 
            // btnSelectBillNameCol
            // 
            this.btnSelectBillNameCol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectBillNameCol.Location = new System.Drawing.Point(323, 186);
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
            this.btnSelectBillProductCol.Location = new System.Drawing.Point(323, 151);
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
            this.txtBillNameColumn.Location = new System.Drawing.Point(90, 187);
            this.txtBillNameColumn.Name = "txtBillNameColumn";
            this.txtBillNameColumn.Size = new System.Drawing.Size(227, 23);
            this.txtBillNameColumn.TabIndex = 9;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(15, 190);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(71, 17);
            this.label6.TabIndex = 8;
            this.label6.Text = "商品名称列:";
            // 
            // txtBillProductColumn
            // 
            this.txtBillProductColumn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtBillProductColumn.Location = new System.Drawing.Point(90, 152);
            this.txtBillProductColumn.Name = "txtBillProductColumn";
            this.txtBillProductColumn.Size = new System.Drawing.Size(227, 23);
            this.txtBillProductColumn.TabIndex = 7;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(15, 155);
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
            this.menuStrip1.Size = new System.Drawing.Size(786, 25);
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
            this.bottomPanel.Location = new System.Drawing.Point(0, 274);
            this.bottomPanel.Name = "bottomPanel";
            this.bottomPanel.Size = new System.Drawing.Size(786, 82);
            this.bottomPanel.TabIndex = 10;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Location = new System.Drawing.Point(0, 356);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(786, 22);
            this.statusStrip1.TabIndex = 11;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.flowLayoutPanel1.Controls.Add(this.btnClose);
            this.flowLayoutPanel1.Controls.Add(this.btnStart);
            this.flowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(12, 41);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(762, 38);
            this.flowLayoutPanel1.TabIndex = 6;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(649, 3);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(110, 32);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnStart
            // 
            this.btnStart.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(123)))), ((int)(((byte)(255)))));
            this.btnStart.FlatAppearance.BorderSize = 0;
            this.btnStart.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnStart.ForeColor = System.Drawing.Color.White;
            this.btnStart.Location = new System.Drawing.Point(533, 3);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(110, 32);
            this.btnStart.TabIndex = 3;
            this.btnStart.Text = "🚀 开始匹配";
            this.btnStart.UseVisualStyleBackColor = false;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // MatchForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(786, 378);
            this.Controls.Add(this.bottomPanel);
            this.Controls.Add(this.gbBill);
            this.Controls.Add(this.gbShipping);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.statusStrip1);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(802, 417);
            this.Name = "MatchForm";
            this.Text = "YY 运单匹配工具 v2.3";
            this.gbShipping.ResumeLayout(false);
            this.gbShipping.PerformLayout();
            this.gbBill.ResumeLayout(false);
            this.gbBill.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.bottomPanel.ResumeLayout(false);
            this.flowLayoutPanel1.ResumeLayout(false);
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
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblShippingInfo;
        private System.Windows.Forms.Label lblBillInfo;
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
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnStart;
    }
}