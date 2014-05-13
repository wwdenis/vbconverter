namespace VBConverter.UI
{
    partial class FileTool
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public FileTool()
        {
            InitializeComponent();
        }

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
            this.gbxFiles = new System.Windows.Forms.GroupBox();
            this.btnReason = new System.Windows.Forms.Button();
            this.btnSelectChange = new System.Windows.Forms.Button();
            this.btnSelectNone = new System.Windows.Forms.Button();
            this.btnSelectAll = new System.Windows.Forms.Button();
            this.lvwFiles = new System.Windows.Forms.ListView();
            this.columnSourceName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnSourceType = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnSourceStatus = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.fbdMain = new System.Windows.Forms.FolderBrowserDialog();
            this.gbxDestination = new System.Windows.Forms.GroupBox();
            this.btnDestinationSelect = new System.Windows.Forms.Button();
            this.txtDestination = new System.Windows.Forms.TextBox();
            this.lblDestination = new System.Windows.Forms.Label();
            this.stsMain = new System.Windows.Forms.StatusStrip();
            this.pbrStatus = new System.Windows.Forms.ToolStripProgressBar();
            this.lblStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.btnConvert = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.gbxSource = new System.Windows.Forms.GroupBox();
            this.btnSourceSelect = new System.Windows.Forms.Button();
            this.txtSource = new System.Windows.Forms.TextBox();
            this.lblSource = new System.Windows.Forms.Label();
            this.cboLanguage = new System.Windows.Forms.ComboBox();
            this.lblLanguage = new System.Windows.Forms.Label();
            this.gbxFiles.SuspendLayout();
            this.gbxDestination.SuspendLayout();
            this.stsMain.SuspendLayout();
            this.gbxSource.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbxFiles
            // 
            this.gbxFiles.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbxFiles.Controls.Add(this.btnReason);
            this.gbxFiles.Controls.Add(this.btnSelectChange);
            this.gbxFiles.Controls.Add(this.btnSelectNone);
            this.gbxFiles.Controls.Add(this.btnSelectAll);
            this.gbxFiles.Controls.Add(this.lvwFiles);
            this.gbxFiles.Location = new System.Drawing.Point(12, 87);
            this.gbxFiles.Name = "gbxFiles";
            this.gbxFiles.Size = new System.Drawing.Size(687, 274);
            this.gbxFiles.TabIndex = 1;
            this.gbxFiles.TabStop = false;
            this.gbxFiles.Text = "Project Files";
            // 
            // btnReason
            // 
            this.btnReason.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnReason.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnReason.Location = new System.Drawing.Point(505, 238);
            this.btnReason.Name = "btnReason";
            this.btnReason.Size = new System.Drawing.Size(170, 26);
            this.btnReason.TabIndex = 4;
            this.btnReason.Text = "Show Convertion Errors";
            this.btnReason.UseVisualStyleBackColor = true;
            this.btnReason.Click += new System.EventHandler(this.ShowErrorReasonHandler);
            // 
            // btnSelectChange
            // 
            this.btnSelectChange.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSelectChange.Location = new System.Drawing.Point(265, 238);
            this.btnSelectChange.Name = "btnSelectChange";
            this.btnSelectChange.Size = new System.Drawing.Size(121, 26);
            this.btnSelectChange.TabIndex = 3;
            this.btnSelectChange.Text = "Invert Selection";
            this.btnSelectChange.UseVisualStyleBackColor = true;
            this.btnSelectChange.Click += new System.EventHandler(this.SelectItemsHandler);
            // 
            // btnSelectNone
            // 
            this.btnSelectNone.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSelectNone.Location = new System.Drawing.Point(138, 238);
            this.btnSelectNone.Name = "btnSelectNone";
            this.btnSelectNone.Size = new System.Drawing.Size(121, 26);
            this.btnSelectNone.TabIndex = 2;
            this.btnSelectNone.Text = "Select None";
            this.btnSelectNone.UseVisualStyleBackColor = true;
            this.btnSelectNone.Click += new System.EventHandler(this.SelectItemsHandler);
            // 
            // btnSelectAll
            // 
            this.btnSelectAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSelectAll.Location = new System.Drawing.Point(11, 238);
            this.btnSelectAll.Name = "btnSelectAll";
            this.btnSelectAll.Size = new System.Drawing.Size(121, 26);
            this.btnSelectAll.TabIndex = 1;
            this.btnSelectAll.Text = "Select All";
            this.btnSelectAll.UseVisualStyleBackColor = true;
            this.btnSelectAll.Click += new System.EventHandler(this.SelectItemsHandler);
            // 
            // lvwFiles
            // 
            this.lvwFiles.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lvwFiles.CheckBoxes = true;
            this.lvwFiles.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnSourceName,
            this.columnSourceType,
            this.columnSourceStatus});
            this.lvwFiles.FullRowSelect = true;
            this.lvwFiles.GridLines = true;
            this.lvwFiles.HideSelection = false;
            this.lvwFiles.Location = new System.Drawing.Point(12, 20);
            this.lvwFiles.Name = "lvwFiles";
            this.lvwFiles.Size = new System.Drawing.Size(663, 212);
            this.lvwFiles.TabIndex = 0;
            this.lvwFiles.UseCompatibleStateImageBehavior = false;
            this.lvwFiles.View = System.Windows.Forms.View.Details;
            this.lvwFiles.SelectedIndexChanged += new System.EventHandler(this.lvwFiles_SelectedIndexChanged);
            this.lvwFiles.DoubleClick += new System.EventHandler(this.ShowErrorReasonHandler);
            // 
            // columnSourceName
            // 
            this.columnSourceName.Text = "File Name";
            this.columnSourceName.Width = 337;
            // 
            // columnSourceType
            // 
            this.columnSourceType.Text = "File Type";
            this.columnSourceType.Width = 75;
            // 
            // columnSourceStatus
            // 
            this.columnSourceStatus.Text = "Status";
            this.columnSourceStatus.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.columnSourceStatus.Width = 90;
            // 
            // gbxDestination
            // 
            this.gbxDestination.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbxDestination.Controls.Add(this.btnDestinationSelect);
            this.gbxDestination.Controls.Add(this.txtDestination);
            this.gbxDestination.Controls.Add(this.lblDestination);
            this.gbxDestination.Location = new System.Drawing.Point(12, 367);
            this.gbxDestination.Name = "gbxDestination";
            this.gbxDestination.Size = new System.Drawing.Size(687, 69);
            this.gbxDestination.TabIndex = 2;
            this.gbxDestination.TabStop = false;
            this.gbxDestination.Text = "Destination Directory";
            // 
            // btnDestinationSelect
            // 
            this.btnDestinationSelect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDestinationSelect.Location = new System.Drawing.Point(574, 32);
            this.btnDestinationSelect.Name = "btnDestinationSelect";
            this.btnDestinationSelect.Size = new System.Drawing.Size(101, 24);
            this.btnDestinationSelect.TabIndex = 2;
            this.btnDestinationSelect.Text = "Browse...";
            this.btnDestinationSelect.UseVisualStyleBackColor = true;
            this.btnDestinationSelect.Click += new System.EventHandler(this.FolderSelectionHandler);
            // 
            // txtDestination
            // 
            this.txtDestination.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDestination.Location = new System.Drawing.Point(12, 34);
            this.txtDestination.Name = "txtDestination";
            this.txtDestination.Size = new System.Drawing.Size(556, 21);
            this.txtDestination.TabIndex = 1;
            this.txtDestination.Validating += new System.ComponentModel.CancelEventHandler(this.FolderValidationHandler);
            // 
            // lblDestination
            // 
            this.lblDestination.AutoSize = true;
            this.lblDestination.Location = new System.Drawing.Point(13, 18);
            this.lblDestination.Name = "lblDestination";
            this.lblDestination.Size = new System.Drawing.Size(29, 13);
            this.lblDestination.TabIndex = 0;
            this.lblDestination.Text = "Path";
            // 
            // stsMain
            // 
            this.stsMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.pbrStatus,
            this.lblStatus});
            this.stsMain.Location = new System.Drawing.Point(0, 484);
            this.stsMain.Name = "stsMain";
            this.stsMain.Size = new System.Drawing.Size(712, 22);
            this.stsMain.TabIndex = 7;
            this.stsMain.Text = "statusStrip1";
            // 
            // pbrStatus
            // 
            this.pbrStatus.Name = "pbrStatus";
            this.pbrStatus.Size = new System.Drawing.Size(100, 16);
            // 
            // lblStatus
            // 
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 17);
            // 
            // btnConvert
            // 
            this.btnConvert.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnConvert.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnConvert.Location = new System.Drawing.Point(12, 445);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(120, 30);
            this.btnConvert.TabIndex = 3;
            this.btnConvert.Text = "Convert";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.CausesValidation = false;
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(579, 445);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(120, 30);
            this.btnClose.TabIndex = 6;
            this.btnClose.Text = "Exit";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // gbxSource
            // 
            this.gbxSource.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbxSource.Controls.Add(this.btnSourceSelect);
            this.gbxSource.Controls.Add(this.txtSource);
            this.gbxSource.Controls.Add(this.lblSource);
            this.gbxSource.Location = new System.Drawing.Point(12, 12);
            this.gbxSource.Name = "gbxSource";
            this.gbxSource.Size = new System.Drawing.Size(687, 69);
            this.gbxSource.TabIndex = 0;
            this.gbxSource.TabStop = false;
            this.gbxSource.Text = "Source Directory";
            // 
            // btnSourceSelect
            // 
            this.btnSourceSelect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSourceSelect.Location = new System.Drawing.Point(574, 31);
            this.btnSourceSelect.Name = "btnSourceSelect";
            this.btnSourceSelect.Size = new System.Drawing.Size(101, 24);
            this.btnSourceSelect.TabIndex = 2;
            this.btnSourceSelect.Text = "Browse...";
            this.btnSourceSelect.UseVisualStyleBackColor = true;
            this.btnSourceSelect.Click += new System.EventHandler(this.FolderSelectionHandler);
            // 
            // txtSource
            // 
            this.txtSource.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSource.Location = new System.Drawing.Point(12, 33);
            this.txtSource.Name = "txtSource";
            this.txtSource.Size = new System.Drawing.Size(556, 21);
            this.txtSource.TabIndex = 1;
            this.txtSource.Validating += new System.ComponentModel.CancelEventHandler(this.FolderValidationHandler);
            // 
            // lblSource
            // 
            this.lblSource.AutoSize = true;
            this.lblSource.Location = new System.Drawing.Point(13, 17);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(29, 13);
            this.lblSource.TabIndex = 0;
            this.lblSource.Text = "Path";
            // 
            // cboLanguage
            // 
            this.cboLanguage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cboLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboLanguage.FormattingEnabled = true;
            this.cboLanguage.Items.AddRange(new object[] {
            "C#",
            "VB .NET"});
            this.cboLanguage.Location = new System.Drawing.Point(174, 451);
            this.cboLanguage.Name = "cboLanguage";
            this.cboLanguage.Size = new System.Drawing.Size(97, 21);
            this.cboLanguage.TabIndex = 5;
            // 
            // lblLanguage
            // 
            this.lblLanguage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblLanguage.AutoSize = true;
            this.lblLanguage.Location = new System.Drawing.Point(139, 455);
            this.lblLanguage.Name = "lblLanguage";
            this.lblLanguage.Size = new System.Drawing.Size(19, 13);
            this.lblLanguage.TabIndex = 4;
            this.lblLanguage.Text = "To";
            // 
            // FileTool
            // 
            this.CancelButton = this.btnClose;
            this.ClientSize = new System.Drawing.Size(712, 506);
            this.Controls.Add(this.cboLanguage);
            this.Controls.Add(this.lblLanguage);
            this.Controls.Add(this.gbxSource);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnConvert);
            this.Controls.Add(this.stsMain);
            this.Controls.Add(this.gbxDestination);
            this.Controls.Add(this.gbxFiles);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "FileTool";
            this.Text = "Batch Converter";
            this.Load += new System.EventHandler(this.FileTool_Load);
            this.gbxFiles.ResumeLayout(false);
            this.gbxDestination.ResumeLayout(false);
            this.gbxDestination.PerformLayout();
            this.stsMain.ResumeLayout(false);
            this.stsMain.PerformLayout();
            this.gbxSource.ResumeLayout(false);
            this.gbxSource.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox gbxFiles;
        private System.Windows.Forms.ListView lvwFiles;
        private System.Windows.Forms.FolderBrowserDialog fbdMain;
        private System.Windows.Forms.Button btnSelectChange;
        private System.Windows.Forms.Button btnSelectNone;
        private System.Windows.Forms.Button btnSelectAll;
        private System.Windows.Forms.GroupBox gbxDestination;
        private System.Windows.Forms.Button btnDestinationSelect;
        private System.Windows.Forms.TextBox txtDestination;
        private System.Windows.Forms.Label lblDestination;
        private System.Windows.Forms.ColumnHeader columnSourceStatus;
        private System.Windows.Forms.ColumnHeader columnSourceName;
        private System.Windows.Forms.ColumnHeader columnSourceType;
        private System.Windows.Forms.StatusStrip stsMain;
        private System.Windows.Forms.ToolStripStatusLabel lblStatus;
        private System.Windows.Forms.ToolStripProgressBar pbrStatus;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.GroupBox gbxSource;
        private System.Windows.Forms.Button btnSourceSelect;
        private System.Windows.Forms.TextBox txtSource;
        private System.Windows.Forms.Label lblSource;
        private System.Windows.Forms.Button btnReason;
        private System.Windows.Forms.ComboBox cboLanguage;
        private System.Windows.Forms.Label lblLanguage;

    }
}