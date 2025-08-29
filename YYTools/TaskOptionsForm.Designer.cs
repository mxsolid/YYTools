namespace YYTools
{
    partial class TaskOptionsForm
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
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.tabControlSettings = new System.Windows.Forms.TabControl();
            this.tabConcatenation = new System.Windows.Forms.TabPage();
            this.cmbSort = new System.Windows.Forms.ComboBox();
            this.lblSort = new System.Windows.Forms.Label();
            this.chkRemoveDuplicates = new System.Windows.Forms.CheckBox();
            this.cmbDelimiter = new System.Windows.Forms.ComboBox();
            this.lblDelimiter = new System.Windows.Forms.Label();
            this.tabPerformance = new System.Windows.Forms.TabPage();
            this.cmbPreviewRows = new System.Windows.Forms.ComboBox();
            this.lblPreviewRows = new System.Windows.Forms.Label();
            this.chkEnableProgressReporting = new System.Windows.Forms.CheckBox();
            this.numMaxPreviewRows = new System.Windows.Forms.NumericUpDown();
            this.lblMaxPreviewRows = new System.Windows.Forms.Label();
            this.numBatchSize = new System.Windows.Forms.NumericUpDown();
            this.lblBatchSize = new System.Windows.Forms.Label();
            this.tabSmartMatching = new System.Windows.Forms.TabPage();
            this.lblMinMatchScoreValue = new System.Windows.Forms.Label();
            this.trkMinMatchScore = new System.Windows.Forms.TrackBar();
            this.lblMinMatchScore = new System.Windows.Forms.Label();
            this.chkEnableExactMatchPriority = new System.Windows.Forms.CheckBox();
            this.chkEnableSmartMatching = new System.Windows.Forms.CheckBox();
            this.tabControlSettings.SuspendLayout();
            this.tabConcatenation.SuspendLayout();
            this.tabPerformance.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numMaxPreviewRows)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numBatchSize)).BeginInit();
            this.tabSmartMatching.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.trkMinMatchScore)).BeginInit();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.Location = new System.Drawing.Point(236, 313);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 25);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "确定";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(317, 313);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 25);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnReset
            // 
            this.btnReset.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnReset.Location = new System.Drawing.Point(12, 313);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(85, 25);
            this.btnReset.TabIndex = 3;
            this.btnReset.Text = "恢复默认";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // tabControlSettings
            // 
            this.tabControlSettings.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControlSettings.Controls.Add(this.tabConcatenation);
            this.tabControlSettings.Controls.Add(this.tabPerformance);
            this.tabControlSettings.Controls.Add(this.tabSmartMatching);
            this.tabControlSettings.Location = new System.Drawing.Point(12, 12);
            this.tabControlSettings.Name = "tabControlSettings";
            this.tabControlSettings.SelectedIndex = 0;
            this.tabControlSettings.Size = new System.Drawing.Size(380, 286);
            this.tabControlSettings.TabIndex = 0;
            // 
            // tabConcatenation
            // 
            this.tabConcatenation.Controls.Add(this.cmbSort);
            this.tabConcatenation.Controls.Add(this.lblSort);
            this.tabConcatenation.Controls.Add(this.chkRemoveDuplicates);
            this.tabConcatenation.Controls.Add(this.cmbDelimiter);
            this.tabConcatenation.Controls.Add(this.lblDelimiter);
            this.tabConcatenation.Location = new System.Drawing.Point(4, 26);
            this.tabConcatenation.Name = "tabConcatenation";
            this.tabConcatenation.Padding = new System.Windows.Forms.Padding(10);
            this.tabConcatenation.Size = new System.Drawing.Size(372, 256);
            this.tabConcatenation.TabIndex = 0;
            this.tabConcatenation.Text = "拼接设置";
            this.tabConcatenation.UseVisualStyleBackColor = true;
            // 
            // cmbSort
            // 
            this.cmbSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSort.FormattingEnabled = true;
            this.cmbSort.Location = new System.Drawing.Point(110, 50);
            this.cmbSort.Name = "cmbSort";
            this.cmbSort.Size = new System.Drawing.Size(121, 25);
            this.cmbSort.TabIndex = 2;
            // 
            // lblSort
            // 
            this.lblSort.AutoSize = true;
            this.lblSort.Location = new System.Drawing.Point(15, 53);
            this.lblSort.Name = "lblSort";
            this.lblSort.Size = new System.Drawing.Size(59, 17);
            this.lblSort.TabIndex = 3;
            this.lblSort.Text = "排序方式:";
            // 
            // chkRemoveDuplicates
            // 
            this.chkRemoveDuplicates.AutoSize = true;
            this.chkRemoveDuplicates.Location = new System.Drawing.Point(18, 93);
            this.chkRemoveDuplicates.Name = "chkRemoveDuplicates";
            this.chkRemoveDuplicates.Size = new System.Drawing.Size(123, 21);
            this.chkRemoveDuplicates.TabIndex = 3;
            this.chkRemoveDuplicates.Text = "拼接结果去除重复";
            this.chkRemoveDuplicates.UseVisualStyleBackColor = true;
            // 
            // cmbDelimiter
            // 
            this.cmbDelimiter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDelimiter.FormattingEnabled = true;
            this.cmbDelimiter.Location = new System.Drawing.Point(110, 15);
            this.cmbDelimiter.Name = "cmbDelimiter";
            this.cmbDelimiter.Size = new System.Drawing.Size(121, 25);
            this.cmbDelimiter.TabIndex = 1;
            // 
            // lblDelimiter
            // 
            this.lblDelimiter.AutoSize = true;
            this.lblDelimiter.Location = new System.Drawing.Point(15, 18);
            this.lblDelimiter.Name = "lblDelimiter";
            this.lblDelimiter.Size = new System.Drawing.Size(83, 17);
            this.lblDelimiter.TabIndex = 0;
            this.lblDelimiter.Text = "多项分隔符:";
            // 
            // tabPerformance
            // 
            this.tabPerformance.Controls.Add(this.cmbPreviewRows);
            this.tabPerformance.Controls.Add(this.lblPreviewRows);
            this.tabPerformance.Controls.Add(this.chkEnableProgressReporting);
            this.tabPerformance.Controls.Add(this.numMaxPreviewRows);
            this.tabPerformance.Controls.Add(this.lblMaxPreviewRows);
            this.tabPerformance.Controls.Add(this.numBatchSize);
            this.tabPerformance.Controls.Add(this.lblBatchSize);
            this.tabPerformance.Location = new System.Drawing.Point(4, 26);
            this.tabPerformance.Name = "tabPerformance";
            this.tabPerformance.Padding = new System.Windows.Forms.Padding(10);
            this.tabPerformance.Size = new System.Drawing.Size(372, 256);
            this.tabPerformance.TabIndex = 1;
            this.tabPerformance.Text = "性能与预览";
            this.tabPerformance.UseVisualStyleBackColor = true;
            // 
            // cmbPreviewRows
            // 
            this.cmbPreviewRows.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPreviewRows.FormattingEnabled = true;
            this.cmbPreviewRows.Location = new System.Drawing.Point(135, 85);
            this.cmbPreviewRows.Name = "cmbPreviewRows";
            this.cmbPreviewRows.Size = new System.Drawing.Size(120, 25);
            this.cmbPreviewRows.TabIndex = 3;
            // 
            // lblPreviewRows
            // 
            this.lblPreviewRows.AutoSize = true;
            this.lblPreviewRows.Location = new System.Drawing.Point(15, 88);
            this.lblPreviewRows.Name = "lblPreviewRows";
            this.lblPreviewRows.Size = new System.Drawing.Size(111, 17);
            this.lblPreviewRows.TabIndex = 5;
            this.lblPreviewRows.Text = "列名解析预览行数:";
            // 
            // chkEnableProgressReporting
            // 
            this.chkEnableProgressReporting.AutoSize = true;
            this.chkEnableProgressReporting.Location = new System.Drawing.Point(18, 128);
            this.chkEnableProgressReporting.Name = "chkEnableProgressReporting";
            this.chkEnableProgressReporting.Size = new System.Drawing.Size(123, 21);
            this.chkEnableProgressReporting.TabIndex = 4;
            this.chkEnableProgressReporting.Text = "启用实时进度报告";
            this.chkEnableProgressReporting.UseVisualStyleBackColor = true;
            // 
            // numMaxPreviewRows
            // 
            this.numMaxPreviewRows.Increment = new decimal(new int[] { 50, 0, 0, 0 });
            this.numMaxPreviewRows.Location = new System.Drawing.Point(135, 50);
            this.numMaxPreviewRows.Maximum = new decimal(new int[] { 1000, 0, 0, 0 });
            this.numMaxPreviewRows.Minimum = new decimal(new int[] { 50, 0, 0, 0 });
            this.numMaxPreviewRows.Name = "numMaxPreviewRows";
            this.numMaxPreviewRows.Size = new System.Drawing.Size(120, 23);
            this.numMaxPreviewRows.TabIndex = 2;
            this.numMaxPreviewRows.Value = new decimal(new int[] { 200, 0, 0, 0 });
            // 
            // lblMaxPreviewRows
            // 
            this.lblMaxPreviewRows.AutoSize = true;
            this.lblMaxPreviewRows.Location = new System.Drawing.Point(15, 52);
            this.lblMaxPreviewRows.Name = "lblMaxPreviewRows";
            this.lblMaxPreviewRows.Size = new System.Drawing.Size(111, 17);
            this.lblMaxPreviewRows.TabIndex = 2;
            this.lblMaxPreviewRows.Text = "数据预览最大行数:";
            // 
            // numBatchSize
            // 
            this.numBatchSize.Increment = new decimal(new int[] { 1000, 0, 0, 0 });
            this.numBatchSize.Location = new System.Drawing.Point(135, 15);
            this.numBatchSize.Maximum = new decimal(new int[] { 50000, 0, 0, 0 });
            this.numBatchSize.Minimum = new decimal(new int[] { 1000, 0, 0, 0 });
            this.numBatchSize.Name = "numBatchSize";
            this.numBatchSize.Size = new System.Drawing.Size(120, 23);
            this.numBatchSize.TabIndex = 1;
            this.numBatchSize.Value = new decimal(new int[] { 5000, 0, 0, 0 });
            // 
            // lblBatchSize
            // 
            this.lblBatchSize.AutoSize = true;
            this.lblBatchSize.Location = new System.Drawing.Point(15, 17);
            this.lblBatchSize.Name = "lblBatchSize";
            this.lblBatchSize.Size = new System.Drawing.Size(95, 17);
            this.lblBatchSize.TabIndex = 0;
            this.lblBatchSize.Text = "数据批处理大小:";
            // 
            // tabSmartMatching
            // 
            this.tabSmartMatching.Controls.Add(this.lblMinMatchScoreValue);
            this.tabSmartMatching.Controls.Add(this.trkMinMatchScore);
            this.tabSmartMatching.Controls.Add(this.lblMinMatchScore);
            this.tabSmartMatching.Controls.Add(this.chkEnableExactMatchPriority);
            this.tabSmartMatching.Controls.Add(this.chkEnableSmartMatching);
            this.tabSmartMatching.Location = new System.Drawing.Point(4, 26);
            this.tabSmartMatching.Name = "tabSmartMatching";
            this.tabSmartMatching.Padding = new System.Windows.Forms.Padding(10);
            this.tabSmartMatching.Size = new System.Drawing.Size(372, 256);
            this.tabSmartMatching.TabIndex = 2;
            this.tabSmartMatching.Text = "智能匹配";
            this.tabSmartMatching.UseVisualStyleBackColor = true;
            // 
            // lblMinMatchScoreValue
            // 
            this.lblMinMatchScoreValue.AutoSize = true;
            this.lblMinMatchScoreValue.Location = new System.Drawing.Point(325, 90);
            this.lblMinMatchScoreValue.Name = "lblMinMatchScoreValue";
            this.lblMinMatchScoreValue.Size = new System.Drawing.Size(31, 17);
            this.lblMinMatchScoreValue.TabIndex = 4;
            this.lblMinMatchScoreValue.Text = "0.80";
            // 
            // trkMinMatchScore
            // 
            this.trkMinMatchScore.Location = new System.Drawing.Point(125, 85);
            this.trkMinMatchScore.Maximum = 100;
            this.trkMinMatchScore.Name = "trkMinMatchScore";
            this.trkMinMatchScore.Size = new System.Drawing.Size(194, 45);
            this.trkMinMatchScore.TabIndex = 3;
            this.trkMinMatchScore.TickFrequency = 10;
            this.trkMinMatchScore.Value = 80;
            this.trkMinMatchScore.ValueChanged += new System.EventHandler(this.trkMinMatchScore_ValueChanged);
            // 
            // lblMinMatchScore
            // 
            this.lblMinMatchScore.AutoSize = true;
            this.lblMinMatchScore.Location = new System.Drawing.Point(15, 90);
            this.lblMinMatchScore.Name = "lblMinMatchScore";
            this.lblMinMatchScore.Size = new System.Drawing.Size(107, 17);
            this.lblMinMatchScore.TabIndex = 2;
            this.lblMinMatchScore.Text = "模糊匹配相似度阈值:";
            // 
            // chkEnableExactMatchPriority
            // 
            this.chkEnableExactMatchPriority.AutoSize = true;
            this.chkEnableExactMatchPriority.Location = new System.Drawing.Point(18, 53);
            this.chkEnableExactMatchPriority.Name = "chkEnableExactMatchPriority";
            this.chkEnableExactMatchPriority.Size = new System.Drawing.Size(147, 21);
            this.chkEnableExactMatchPriority.TabIndex = 2;
            this.chkEnableExactMatchPriority.Text = "优先使用完全匹配结果";
            this.chkEnableExactMatchPriority.UseVisualStyleBackColor = true;
            // 
            // chkEnableSmartMatching
            // 
            this.chkEnableSmartMatching.AutoSize = true;
            this.chkEnableSmartMatching.Location = new System.Drawing.Point(18, 18);
            this.chkEnableSmartMatching.Name = "chkEnableSmartMatching";
            this.chkEnableSmartMatching.Size = new System.Drawing.Size(147, 21);
            this.chkEnableSmartMatching.TabIndex = 1;
            this.chkEnableSmartMatching.Text = "启用智能模糊匹配模式";
            this.chkEnableSmartMatching.UseVisualStyleBackColor = true;
            // 
            // TaskOptionsForm
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(404, 351);
            this.Controls.Add(this.tabControlSettings);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(420, 390);
            this.Name = "TaskOptionsForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "任务选项配置";
            this.tabControlSettings.ResumeLayout(false);
            this.tabConcatenation.ResumeLayout(false);
            this.tabConcatenation.PerformLayout();
            this.tabPerformance.ResumeLayout(false);
            this.tabPerformance.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numMaxPreviewRows)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numBatchSize)).EndInit();
            this.tabSmartMatching.ResumeLayout(false);
            this.tabSmartMatching.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.trkMinMatchScore)).EndInit();
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.TabControl tabControlSettings;
        private System.Windows.Forms.TabPage tabConcatenation;
        private System.Windows.Forms.TabPage tabPerformance;
        private System.Windows.Forms.TabPage tabSmartMatching;
        private System.Windows.Forms.Label lblDelimiter;
        private System.Windows.Forms.ComboBox cmbDelimiter;
        private System.Windows.Forms.CheckBox chkRemoveDuplicates;
        private System.Windows.Forms.ComboBox cmbSort;
        private System.Windows.Forms.Label lblSort;
        private System.Windows.Forms.Label lblBatchSize;
        private System.Windows.Forms.NumericUpDown numBatchSize;
        private System.Windows.Forms.Label lblMaxPreviewRows;
        private System.Windows.Forms.NumericUpDown numMaxPreviewRows;
        private System.Windows.Forms.CheckBox chkEnableProgressReporting;
        private System.Windows.Forms.Label lblPreviewRows;
        private System.Windows.Forms.ComboBox cmbPreviewRows;
        private System.Windows.Forms.CheckBox chkEnableSmartMatching;
        private System.Windows.Forms.CheckBox chkEnableExactMatchPriority;
        private System.Windows.Forms.Label lblMinMatchScore;
        private System.Windows.Forms.TrackBar trkMinMatchScore;
        private System.Windows.Forms.Label lblMinMatchScoreValue;
    }
}