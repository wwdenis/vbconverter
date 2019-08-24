using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using VBConverter.CodeParser;
using VBConverter.CodeWriter;

namespace VBConverter.UI
{
    public partial class FileTool : Form
    {
        private enum SelectMode { All, None, Change }

        private bool ValidatePath(string path, bool ignoreEmpty)
        {
            bool exists = (ignoreEmpty && string.IsNullOrEmpty(path)) || Directory.Exists(path);

            if (!exists)
                MessageBox.Show("Invalid folder !", "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            return exists;
        }

        private void SelectItems(SelectMode mode)
        {
            foreach (ListViewItem item in lvwFiles.Items)
            {
                switch (mode)
                {
                    case SelectMode.All:
                        item.Checked = true;
                        break;
                    case SelectMode.None:
                        item.Checked = false;
                        break;
                    case SelectMode.Change:
                        item.Checked = !item.Checked;
                        break;
                }
            }
        }

        private void ListFiles()
        {
            string path = txtSource.Text;

            lvwFiles.Items.Clear();

            if (string.IsNullOrEmpty(path) || !Directory.Exists(path))
                return;

            Dictionary<string, string> types = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            types.Add(".bas", "Module");
            types.Add(".cls", "Class");
            types.Add(".frm", "Form");

            DirectoryInfo dir = new DirectoryInfo(path);

            foreach (FileInfo file in dir.GetFiles("*.*", SearchOption.AllDirectories))
            {
                if (types.ContainsKey(file.Extension))
                {
                    string type = types[file.Extension];
                    string name = file.FullName;
                    name = name.Remove(0, path.Length);

                    ListViewItem item = new ListViewItem(new string[] { name, type, string.Empty });
                    item.UseItemStyleForSubItems = false;
                    item.Checked = true;
                    lvwFiles.Items.Add(item);
                }
            }
            bool hasItems = lvwFiles.Items.Count > 0;
            gbxFiles.Enabled = hasItems;
            gbxDestination.Enabled = hasItems;
            btnConvert.Enabled = hasItems;
        }

        private void FolderValidationHandler(object sender, CancelEventArgs e)
        {
            TextBox TextBoxObject = null;

            try
            {
                TextBoxObject = (TextBox)sender;
                string path = TextBoxObject.Text;
                bool valid = ValidatePath(path, true);

                if (!valid)
                {
                    e.Cancel = true;
                    TextBoxObject.Focus();
                }
                else if (path.EndsWith(@"\"))
                {
                    path = path.Remove(path.Length - 1);
                }

                while (path.Contains(@"\\"))
                    path = path.Replace(@"\\", @"\");

                TextBoxObject.Text = path;

                if (object.ReferenceEquals(sender, txtSource))
                    ListFiles();
            }
            catch (Exception ex)
            {
                if (TextBoxObject !=null)
                    TextBoxObject.Text = string.Empty;
                MessageBox.Show("There's a error trying to validate the directory. \n\nDetails: " + ex.Message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FolderSelectionHandler(object sender, EventArgs e)
        {
            try
            {
                TextBox TextBoxObject = null;

                if (object.ReferenceEquals(sender, btnSourceSelect))
                    TextBoxObject = txtSource;
                else if (object.ReferenceEquals(sender, btnDestinationSelect))
                    TextBoxObject = txtDestination;
                else
                    return;

                if (Directory.Exists(TextBoxObject.Text))
                    fbdMain.SelectedPath = TextBoxObject.Text;

                DialogResult result = fbdMain.ShowDialog();

                if (result == DialogResult.OK)
                {
                    string path = fbdMain.SelectedPath;
                    if (path.EndsWith(@"\"))
                        path.Remove(TextBoxObject.Text.Length - 1);
                    TextBoxObject.Text = path;

                    if (object.ReferenceEquals(sender, btnSourceSelect))
                        ListFiles();
                }

                TextBoxObject.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("There's a error trying to select a directory. \n\nDetails: " + ex.Message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SelectItemsHandler(object sender, EventArgs e)
        {
            try
            {
                if (object.ReferenceEquals(sender, btnSelectAll))
                    SelectItems(SelectMode.All);
                else if (object.ReferenceEquals(sender, btnSelectNone))
                    SelectItems(SelectMode.None);
                else if (object.ReferenceEquals(sender, btnSelectChange))
                    SelectItems(SelectMode.Change);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Não é possível processar a requisição. \n\nDetails: " + ex.Message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void FileTool_Load(object sender, EventArgs e)
        {
            try
            {
                gbxFiles.Enabled = false;
                gbxDestination.Enabled = false;
                btnConvert.Enabled = false;
                pbrStatus.Visible = false;
                btnReason.Visible = false;
                cboLanguage.Text = "C#";
            }
            catch (Exception ex)
            {
                MessageBox.Show("There's a error trying to open the window. \n\nDetails: " + ex.Message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(txtDestination.Text))
                {
                    MessageBox.Show("A pasta de destino não foi informado. \n\nInforma um caminho válido para a gravação do resultado da conversão.", "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtDestination.Focus();
                    return;
                }

                int fileCount = lvwFiles.CheckedItems.Count;

                if (fileCount == 0)
                {
                    MessageBox.Show("Não há arquivos selecionados para conversão. \n\nSelecione os arquivos clicando na caixa de checagem à esquerda do nome.", "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    lvwFiles.Focus();
                    return;
                }

                this.Cursor = Cursors.WaitCursor;
                lblStatus.Text = "Iniciando...";
                pbrStatus.Maximum = fileCount;
                pbrStatus.Value = 0;

                pbrStatus.Visible = true;
                lblStatus.Visible = true;

                foreach (ListViewItem item in lvwFiles.CheckedItems)
                {
                    item.SubItems[2].Text = string.Empty;
                    item.SubItems[2].ForeColor = item.ForeColor;
                    item.SubItems[2].Font = new Font(item.Font, FontStyle.Regular);
                    item.Tag = string.Empty;
                }

                Application.DoEvents();
                int sucessCount = 0;
                int errorCount = 0;

                foreach (ListViewItem item in lvwFiles.CheckedItems)
                {
                    pbrStatus.Value++;
                    lblStatus.Text = "Convertendo o arquivo " + item.Text;
                    item.SubItems[2].Text = "Convertendo...";
                    item.EnsureVisible();

                    string reason = string.Empty;
                    bool success = ConvertFile(item.Text, out reason);
                    if (success)
                        sucessCount++;
                    else
                        errorCount++;

                    item.SubItems[2].Text = success ? "Sucesso" : "Erro";
                    item.SubItems[2].ForeColor = success ? Color.Blue : Color.Red;
                    item.SubItems[2].Font = new Font(item.SubItems[2].Font, FontStyle.Bold);
                    item.Tag = reason;
                }

                string message = "Batch conversion sucessfully done !";
                if (errorCount > 0)
                    message = string.Format("The batch conversion has finished, but {0} files couldn't be converted !\n\nCheck the conversion errors by pressing the button 'Show Conversion Erros'.", errorCount);

                MessageBox.Show(message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("There's a error trying to convert!. \n\nDetalhes: " + ex.Message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                lblStatus.Visible = false;
                pbrStatus.Visible = false;
                this.Cursor = Cursors.Default;
            }
        }

        private bool ConvertFile(string file, out string errorReason)
        {
            try
            {
                Application.DoEvents();

                string sourcePath = txtSource.Text + file;
                string sourceText = File.ReadAllText(sourcePath);

                FileInfo info = new FileInfo(sourcePath);
                string fileName = info.Name;
                fileName = fileName.Remove(fileName.Length - info.Extension.Length, info.Extension.Length);

                sourceText = sourceText.Trim();
                if (string.IsNullOrEmpty(sourceText))
                {
                    errorReason = string.Format("The file {0} is empty !", sourcePath);
                    return false;
                }

                ConverterEngine engine = new ConverterEngine(LanguageVersion.VB6, sourceText);
                if (cboLanguage.Text == "VB .NET")
                    engine.ResultType = DestinationLanguage.VisualBasic;
                engine.FileName = fileName;
                bool success = engine.Convert();

                if (success)
                {
                    string resultPath = string.Format(@"{0}\{1}.{2}", txtDestination.Text, fileName, engine.ResultFileExtension);
                    string resultText = engine.Result;
                    File.WriteAllText(resultPath, resultText);
                }

                errorReason = engine.GetErrors();
                return success;
            }
            catch (Exception ex)
            {
                errorReason = string.Format("Cannot convert due to internal error.\n\nDetails: " + ex.Message);
                return false;
            }
        }

        private void ShowErrorReasonHandler(object sender, EventArgs e)
        {
            try
            {
                string text = string.Empty;
                foreach (ListViewItem item in lvwFiles.SelectedItems)
                {
                    if (!string.IsNullOrEmpty(text))
                        text += '\n';
                    text += item.Tag;
                }

                if (!string.IsNullOrEmpty(text))
                    MessageBox.Show("This file cannot be converted.\n\nDetails: " + text, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {
                MessageBox.Show("This file cannot be converted.\n\nDetails: " + ex.Message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void lvwFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                foreach (ListViewItem item in lvwFiles.SelectedItems)
                    btnReason.Visible = !string.IsNullOrEmpty(item.Tag as string);
            }
            catch (Exception ex)
            {
                MessageBox.Show("There's a error trying to select a item.\n\nDetails: " + ex.Message, "VB Converter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}