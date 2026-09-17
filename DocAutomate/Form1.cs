using DocAutomate.Services;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace DocAutomate
{
    public partial class Form1 : Form
    {
        private readonly ExcelDocumentService documentService = new ExcelDocumentService();

        public Form1()
        {
            InitializeComponent();
            selectedDateCalendar.MinDate = DateTime.Today;
        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select an Excel file";
                dialog.Filter = "Excel workbooks (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls";
                dialog.DefaultExt = "xlsx";
                dialog.CheckFileExists = true;
                dialog.CheckPathExists = true;
                dialog.RestoreDirectory = true;

                if (dialog.ShowDialog(this) == DialogResult.OK)
                    sourceFileTextBox.Text = dialog.FileName;
            }
        }

        private void powerpointBrowseButton_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select the Software changes PowerPoint";
                dialog.Filter = "PowerPoint presentations (*.pptx)|*.pptx";
                dialog.CheckFileExists = true;
                dialog.RestoreDirectory = true;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    powerpointTextBox.Text = dialog.FileName;
            }
        }

        private void generateButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(sourceFileTextBox.Text))
            {
                ShowWarning("Please select an Excel file first.", "Select a file");
                return;
            }

            var nameInputs = new[] { firstNameTextBox, secondNameTextBox };
            var nameParts = new string[nameInputs.Length];
            for (int i = 0; i < nameInputs.Length; i++)
            {
                string part = nameInputs[i].Text.Trim();
                string extension = Path.GetExtension(part);
                if (ExcelDocumentService.IsSupportedExtension(extension))
                    part = part.Substring(0, part.Length - extension.Length);

                if (!ExcelDocumentService.IsValidNamePart(part))
                {
                    ShowWarning("Please enter a valid name in " + (i == 0 ? "Name 1" : "Name 2") + ".",
                        "Invalid file name");
                    nameInputs[i].Focus();
                    return;
                }

                nameParts[i] = part;
            }

            selectedDateCalendar.MinDate = DateTime.Today;
            if (selectedDateCalendar.SelectionStart.Date < DateTime.Today)
            {
                ShowWarning("Please select today or a future date.", "Invalid date");
                return;
            }

            try
            {
                string destination = documentService.GetDestination(sourceFileTextBox.Text, nameParts, selectedDateCalendar.SelectionStart);
                if (File.Exists(destination) || Directory.Exists(destination))
                {
                    ShowWarning("A file or folder already exists at:\n" + destination
                        + "\nPlease enter a different name.", "Name already exists");
                    return;
                }

                documentService.Generate(sourceFileTextBox.Text, destination, powerpointTextBox.Text.Trim());
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "explorer.exe"),
                        Arguments = "/select,\"" + destination + "\"",
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    ShowWarning("Excel document generated:\n" + destination
                        + "\n\nCould not open its folder.\n" + ex.Message, "Document generated");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Could not generate the Excel document.\n" + ex.Message,
                    "Generate failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowWarning(string message, string title)
        {
            MessageBox.Show(this, message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}

