using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Fireball.CodeEditor.SyntaxFiles;
using VBConverter.CodeParser;
using VBConverter.CodeWriter;

namespace VBConverter.UI
{
    public partial class CodeTool : Form
    {
        private void CodeTool_Load(object sender, EventArgs e)
        {
            try
            {
                CodeEditorSyntaxLoader.SetSyntax(codeEditorSource, SyntaxLanguage.VB);
                cboLanguage.Text = "C#";
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                MessageBox.Show("It's not possible to open the main window. \n\nDetails: " + ex.Message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CodeTool_Resize(object sender, EventArgs e)
        {
            try
            {
                gbxSource.Height = (this.Height - 102) / 2;
                gbxDestination.Height = gbxSource.Height;
                gbxDestination.Top = gbxSource.Top + gbxSource.Height + 6;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(string.Format("Error while resizing the window. Details: {0}", ex));
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(codeEditorSource.Document.Text))
                {
                    MessageBox.Show("There's no VB6 code to convert !", "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    codeEditorSource.Focus();
                }
                else
                {
                    this.Cursor = Cursors.WaitCursor;

                    string source = codeEditorSource.Document.Text;
                    ConverterEngine engine = new ConverterEngine(LanguageVersion.VB6, source);
                    if (cboLanguage.Text == "VB .NET")
                        engine.ResultType = DestinationLanguage.VisualBasic;
                    bool success = engine.Convert();
                    string result = success ? engine.Result : engine.GetErrors();
                    codeEditorDestination.Document.Text = result;
                    codeEditorDestination.Focus();

                    if (engine.Errors.Count > 0)
                        MessageBox.Show("It's not possible to convert the the source code due to compile errors!\n\nCheck the sintax.", "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                MessageBox.Show("It's not possible to convert the source code. \n\nDetails: " + ex.Message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            try
            {
                openFileSource.Filter = "Basic Files (*.bas)|*.bas|Class Files (*.cls)|*.cls|Form Files (*.frm)|*.frm|All files (*.*)|*.*";
                DialogResult result = openFileSource.ShowDialog();
                if (result == DialogResult.OK)
                {
                    string source = File.ReadAllText(openFileSource.FileName, Encoding.Default);
                    codeEditorSource.Document.Text = ConverterEngine.FilterSource(source);
                }
                codeEditorSource.Focus();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                MessageBox.Show("It's not possible to open the dialog. \n\nDetails: " + ex.Message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(codeEditorDestination.Document.Text))
                {
                    MessageBox.Show("There's no source code to save !", "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    saveFileDestination.Filter = "C# Files (*.cs)|*.cs|Text Files (*.txt)|*.frm|All files (*.*)|*.*";
                    DialogResult result = saveFileDestination.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        codeEditorDestination.Save(saveFileDestination.FileName);
                        MessageBox.Show(string.Format("Conversion saved !\n\nCaminho: {0}", saveFileDestination.FileName), "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                
                codeEditorDestination.Focus();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                MessageBox.Show("It's not possible to save! \n\nDetails: " + ex.Message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cboLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                SyntaxLanguage language = SyntaxLanguage.CSharp;
                if (cboLanguage.Text == "VB .NET")
                    language = SyntaxLanguage.VBNET;
                CodeEditorSyntaxLoader.SetSyntax(codeEditorDestination, language);
                codeEditorDestination.Document.Clear();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                MessageBox.Show("It's not possible to select the destination language. \n\nDetails: " + ex.Message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnBatch_Click(object sender, EventArgs e)
        {
            try
            {
                var window = new FileTool
                {
                    StartPosition = FormStartPosition.CenterParent
                };

                window.ShowDialog();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                MessageBox.Show("It's not possible to open the Batch Convertion Window. \n\nDetails: " + ex.Message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}