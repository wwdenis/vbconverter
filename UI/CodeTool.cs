using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
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
                if (string.IsNullOrEmpty(SourceEditor.Text))
                {
                    MessageBox.Show("There's no source code to convert !", "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    SourceEditor.Focus();
                }
                else
                {
                    this.Cursor = Cursors.WaitCursor;

                    string source = SourceEditor.Text;
                    ConverterEngine engine = new ConverterEngine(LanguageVersion.VB6, source);
                    if (cboLanguage.Text == "VB .NET")
                        engine.ResultType = DestinationLanguage.VisualBasic;
                    bool success = engine.Convert();
                    string result = success ? engine.Result : engine.GetErrors();
                    DestinationEditor.Text = result;
                    DestinationEditor.Focus();

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
                    SourceEditor.Text = ConverterEngine.FilterSource(source);
                }
                SourceEditor.Focus();
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
                if (string.IsNullOrEmpty(DestinationEditor.Text))
                {
                    MessageBox.Show("There's no source code to save !", "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    saveFileDestination.Filter = "C# Files (*.cs)|*.cs|Text Files (*.txt)|*.frm|All files (*.*)|*.*";
                    DialogResult result = saveFileDestination.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        File.WriteAllText(saveFileDestination.FileName, DestinationEditor.Text);
                        MessageBox.Show(string.Format("Conversion saved !\n\nCaminho: {0}", saveFileDestination.FileName), "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

                DestinationEditor.Focus();
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
                DestinationEditor.Clear();

                if (cboLanguage.Text == "VB .NET")
                {
                    SyntaxHandler.SetVisualBasic(SourceEditor);
                }
                else
                {
                    SyntaxHandler.SetCSharp(SourceEditor);
                }
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

        private void SourceEditor_TextChanged(object sender, EventArgs e)
        {
            SyntaxHandler.SetVisualBasic(SourceEditor);
        }

        private void DestinationEditor_TextChanged(object sender, EventArgs e)
        {
            SyntaxHandler.SetCSharp(DestinationEditor);
        }
    }
}