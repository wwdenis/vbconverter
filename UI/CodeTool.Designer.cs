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
            this.gbxSource = new System.Windows.Forms.GroupBox();
            this.SourceEditor = new System.Windows.Forms.RichTextBox();
            this.gbxDestination = new System.Windows.Forms.GroupBox();
            this.DestinationEditor = new System.Windows.Forms.RichTextBox();
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
            this.gbxSource.Controls.Add(this.SourceEditor);
            this.gbxSource.Location = new System.Drawing.Point(12, 8);
            this.gbxSource.Name = "gbxSource";
            this.gbxSource.Size = new System.Drawing.Size(688, 222);
            this.gbxSource.TabIndex = 0;
            this.gbxSource.TabStop = false;
            this.gbxSource.Text = "Visual Basic 6";
            // 
            // SourceEditor
            // 
            this.SourceEditor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.SourceEditor.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SourceEditor.Location = new System.Drawing.Point(7, 21);
            this.SourceEditor.Name = "SourceEditor";
            this.SourceEditor.Size = new System.Drawing.Size(675, 195);
            this.SourceEditor.TabIndex = 1;
            this.SourceEditor.Text = "";
            this.SourceEditor.TextChanged += new System.EventHandler(this.SourceEditor_TextChanged);
            // 
            // gbxDestination
            // 
            this.gbxDestination.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbxDestination.Controls.Add(this.DestinationEditor);
            this.gbxDestination.Location = new System.Drawing.Point(12, 236);
            this.gbxDestination.Name = "gbxDestination";
            this.gbxDestination.Size = new System.Drawing.Size(688, 222);
            this.gbxDestination.TabIndex = 2;
            this.gbxDestination.TabStop = false;
            this.gbxDestination.Text = ".NET";
            // 
            // DestinationEditor
            // 
            this.DestinationEditor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DestinationEditor.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DestinationEditor.Location = new System.Drawing.Point(7, 20);
            this.DestinationEditor.Name = "DestinationEditor";
            this.DestinationEditor.Size = new System.Drawing.Size(675, 195);
            this.DestinationEditor.TabIndex = 3;
            this.DestinationEditor.Text = "";
            this.DestinationEditor.TextChanged += new System.EventHandler(this.DestinationEditor_TextChanged);
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
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.OpenFileDialog openFileSource;
        private System.Windows.Forms.SaveFileDialog saveFileDestination;
        private System.Windows.Forms.Label lblLanguage;
        private System.Windows.Forms.ComboBox cboLanguage;
        private System.Windows.Forms.Button btnBatch;
        private System.Windows.Forms.RichTextBox SourceEditor;
        private System.Windows.Forms.RichTextBox DestinationEditor;
    }
}

