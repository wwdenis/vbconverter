namespace VBConverter.UI
{
    partial class CodeTool
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CodeTool()
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
            this.components = new System.ComponentModel.Container();
            Fireball.Windows.Forms.LineMarginRender lineMarginRender1 = new Fireball.Windows.Forms.LineMarginRender();
            this.gbxSource = new System.Windows.Forms.GroupBox();
            this.codeEditorSource = new Fireball.Windows.Forms.CodeEditorControl();
            this.syntaxDocumentSource = new Fireball.Syntax.SyntaxDocument(this.components);
            this.gbxDestination = new System.Windows.Forms.GroupBox();
            this.codeEditorDestination = new Fireball.Windows.Forms.CodeEditorControl();
            this.syntaxDocumentDestination = new Fireball.Syntax.SyntaxDocument(this.components);
            this.btnConvert = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnLoad = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.openFileSource = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDestination = new System.Windows.Forms.SaveFileDialog();
            this.lblLanguage = new System.Windows.Forms.Label();
            this.cboLanguage = new System.Windows.Forms.ComboBox();
            this.btnBatch = new System.Windows.Forms.Button();
            this.gbxSource.SuspendLayout();
            this.gbxDestination.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbxSource
            // 
            this.gbxSource.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbxSource.Controls.Add(this.codeEditorSource);
            this.gbxSource.Location = new System.Drawing.Point(12, 8);
            this.gbxSource.Name = "gbxSource";
            this.gbxSource.Size = new System.Drawing.Size(688, 222);
            this.gbxSource.TabIndex = 0;
            this.gbxSource.TabStop = false;
            this.gbxSource.Text = "Visual Basic 6";
            // 
            // codeEditorSource
            // 
            this.codeEditorSource.ActiveView = Fireball.Windows.Forms.CodeEditor.ActiveView.BottomRight;
            this.codeEditorSource.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.codeEditorSource.AutoListPosition = null;
            this.codeEditorSource.AutoListSelectedText = "a123";
            this.codeEditorSource.AutoListVisible = false;
            this.codeEditorSource.CopyAsRTF = false;
            this.codeEditorSource.Document = this.syntaxDocumentSource;
            this.codeEditorSource.InfoTipCount = 1;
            this.codeEditorSource.InfoTipPosition = null;
            this.codeEditorSource.InfoTipSelectedIndex = 1;
            this.codeEditorSource.InfoTipVisible = false;
            lineMarginRender1.Bounds = new System.Drawing.Rectangle(19, 0, 19, 16);
            this.codeEditorSource.LineMarginRender = lineMarginRender1;
            this.codeEditorSource.Location = new System.Drawing.Point(6, 20);
            this.codeEditorSource.LockCursorUpdate = false;
            this.codeEditorSource.Name = "codeEditorSource";
            this.codeEditorSource.Saved = false;
            this.codeEditorSource.ShowScopeIndicator = false;
            this.codeEditorSource.Size = new System.Drawing.Size(676, 196);
            this.codeEditorSource.SmoothScroll = false;
            this.codeEditorSource.SplitviewH = -4;
            this.codeEditorSource.SplitviewV = -4;
            this.codeEditorSource.TabGuideColor = System.Drawing.Color.FromArgb(((int)(((byte)(244)))), ((int)(((byte)(243)))), ((int)(((byte)(234)))));
            this.codeEditorSource.TabIndex = 1;
            this.codeEditorSource.WhitespaceColor = System.Drawing.SystemColors.ControlDark;
            // 
            // syntaxDocumentSource
            // 
            this.syntaxDocumentSource.Lines = new string[] {
        ""};
            this.syntaxDocumentSource.MaxUndoBufferSize = 1000;
            this.syntaxDocumentSource.Modified = false;
            this.syntaxDocumentSource.UndoStep = 0;
            // 
            // gbxDestination
            // 
            this.gbxDestination.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbxDestination.Controls.Add(this.codeEditorDestination);
            this.gbxDestination.Location = new System.Drawing.Point(12, 236);
            this.gbxDestination.Name = "gbxDestination";
            this.gbxDestination.Size = new System.Drawing.Size(688, 222);
            this.gbxDestination.TabIndex = 2;
            this.gbxDestination.TabStop = false;
            this.gbxDestination.Text = ".NET";
            // 
            // codeEditorDestination
            // 
            this.codeEditorDestination.ActiveView = Fireball.Windows.Forms.CodeEditor.ActiveView.BottomRight;
            this.codeEditorDestination.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.codeEditorDestination.AutoListPosition = null;
            this.codeEditorDestination.AutoListSelectedText = "a123";
            this.codeEditorDestination.AutoListVisible = false;
            this.codeEditorDestination.CopyAsRTF = false;
            this.codeEditorDestination.Document = this.syntaxDocumentDestination;
            this.codeEditorDestination.InfoTipCount = 1;
            this.codeEditorDestination.InfoTipPosition = null;
            this.codeEditorDestination.InfoTipSelectedIndex = 1;
            this.codeEditorDestination.InfoTipVisible = false;
            this.codeEditorDestination.LineMarginRender = lineMarginRender1;
            this.codeEditorDestination.Location = new System.Drawing.Point(6, 20);
            this.codeEditorDestination.LockCursorUpdate = false;
            this.codeEditorDestination.Name = "codeEditorDestination";
            this.codeEditorDestination.Saved = false;
            this.codeEditorDestination.ShowScopeIndicator = false;
            this.codeEditorDestination.Size = new System.Drawing.Size(676, 196);
            this.codeEditorDestination.SmoothScroll = false;
            this.codeEditorDestination.SplitviewH = -4;
            this.codeEditorDestination.SplitviewV = -4;
            this.codeEditorDestination.TabGuideColor = System.Drawing.Color.FromArgb(((int)(((byte)(244)))), ((int)(((byte)(243)))), ((int)(((byte)(234)))));
            this.codeEditorDestination.TabIndex = 3;
            this.codeEditorDestination.WhitespaceColor = System.Drawing.SystemColors.ControlDark;
            // 
            // syntaxDocumentDestination
            // 
            this.syntaxDocumentDestination.Lines = new string[] {
        ""};
            this.syntaxDocumentDestination.MaxUndoBufferSize = 1000;
            this.syntaxDocumentDestination.Modified = false;
            this.syntaxDocumentDestination.UndoStep = 0;
            // 
            // btnConvert
            // 
            this.btnConvert.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnConvert.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnConvert.Location = new System.Drawing.Point(12, 464);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(105, 30);
            this.btnConvert.TabIndex = 4;
            this.btnConvert.Text = "Convert";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.CausesValidation = false;
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(613, 464);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(87, 30);
            this.btnClose.TabIndex = 10;
            this.btnClose.Text = "Exit";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnLoad
            // 
            this.btnLoad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnLoad.Location = new System.Drawing.Point(287, 464);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(101, 30);
            this.btnLoad.TabIndex = 7;
            this.btnLoad.Text = "Load File";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // btnSave
            // 
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.Location = new System.Drawing.Point(394, 464);
            this.btnSave.Name = "btnSave";
            this.btnSave.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btnSave.Size = new System.Drawing.Size(97, 30);
            this.btnSave.TabIndex = 8;
            this.btnSave.Text = "Save Result";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // lblLanguage
            // 
            this.lblLanguage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblLanguage.AutoSize = true;
            this.lblLanguage.Location = new System.Drawing.Point(133, 474);
            this.lblLanguage.Name = "lblLanguage";
            this.lblLanguage.Size = new System.Drawing.Size(19, 13);
            this.lblLanguage.TabIndex = 5;
            this.lblLanguage.Text = "To";
            // 
            // cboLanguage
            // 
            this.cboLanguage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.cboLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboLanguage.FormattingEnabled = true;
            this.cboLanguage.Items.AddRange(new object[] {
            "C#",
            "VB .NET"});
            this.cboLanguage.Location = new System.Drawing.Point(158, 469);
            this.cboLanguage.Name = "cboLanguage";
            this.cboLanguage.Size = new System.Drawing.Size(74, 21);
            this.cboLanguage.TabIndex = 6;
            this.cboLanguage.SelectedIndexChanged += new System.EventHandler(this.cboLanguage_SelectedIndexChanged);
            // 
            // btnBatch
            // 
            this.btnBatch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBatch.Location = new System.Drawing.Point(497, 465);
            this.btnBatch.Name = "btnBatch";
            this.btnBatch.Size = new System.Drawing.Size(110, 30);
            this.btnBatch.TabIndex = 9;
            this.btnBatch.Text = "Batch Converter";
            this.btnBatch.UseVisualStyleBackColor = true;
            this.btnBatch.Click += new System.EventHandler(this.btnBatch_Click);
            // 
            // CodeTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnClose;
            this.ClientSize = new System.Drawing.Size(712, 506);
            this.Controls.Add(this.cboLanguage);
            this.Controls.Add(this.lblLanguage);
            this.Controls.Add(this.btnBatch);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnConvert);
            this.Controls.Add(this.gbxDestination);
            this.Controls.Add(this.gbxSource);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "CodeTool";
            this.Text = "Code Converter";
            this.Load += new System.EventHandler(this.CodeTool_Load);
            this.Resize += new System.EventHandler(this.CodeTool_Resize);
            this.gbxSource.ResumeLayout(false);
            this.gbxDestination.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox gbxSource;
        private System.Windows.Forms.GroupBox gbxDestination;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Button btnClose;
        private Fireball.Windows.Forms.CodeEditorControl codeEditorSource;
        private Fireball.Syntax.SyntaxDocument syntaxDocumentSource;
        private Fireball.Windows.Forms.CodeEditorControl codeEditorDestination;
        private Fireball.Syntax.SyntaxDocument syntaxDocumentDestination;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.OpenFileDialog openFileSource;
        private System.Windows.Forms.SaveFileDialog saveFileDestination;
        private System.Windows.Forms.Label lblLanguage;
        private System.Windows.Forms.ComboBox cboLanguage;
        private System.Windows.Forms.Button btnBatch;


    }
}

