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
        }

        protected override void OnLoad(EventArgs e)
        {
            if (System.ComponentModel.LicenseManager.UsageMode != System.ComponentModel.LicenseUsageMode.Designtime && !DesignMode)
            {
                languageComboBox.SelectedIndex = AppText.IsKorean ? 1 : 0;
                ApplyLanguage();
                selectedDateCalendar.MinDate = DateTime.Today;
            }
            base.OnLoad(e);
        }

        private void languageComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (System.ComponentModel.LicenseManager.UsageMode == System.ComponentModel.LicenseUsageMode.Designtime || DesignMode)
                return;
            AppText.IsKorean = languageComboBox.SelectedIndex == 1;
            ApplyLanguage();
        }

        private void ApplyLanguage()
        {
            SuspendLayout();
            Text = AppText.Get("Excel Document Generator");
            browseButton.Text = powerpointBrowseButton.Text = AppText.Get("Browse");
            generateButton.Text = AppText.Get("Generate");
            sourceFileLabel.Text = AppText.Get("Excel file");
            powerpointLabel.Text = AppText.Get("PowerPoint file");
            firstNameLabel.Text = AppText.Get("Part");
            secondNameLabel.Text = AppText.Get("CRC");
            label1.Text = AppText.Get("Model");
            // Malgun Gothic supports Hangul and Latin text on Windows.
            if (AppText.IsKorean && Font.Name != "Malgun Gothic")
                Font = new System.Drawing.Font("Malgun Gothic", 9F);
            ResumeLayout(true);
        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = AppText.Get("Select an Excel file");
                dialog.Filter = AppText.Get("Excel workbooks (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls");
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
                dialog.Title = AppText.Get("Select the Software changes PowerPoint");
                dialog.Filter = AppText.Get("PowerPoint presentations (*.pptx)|*.pptx");
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
                ShowWarning(AppText.Get("Please select an Excel file first."), AppText.Get("Select a file"));
                return;
            }

            string modelName = textBox1.Text.Trim();
            if (!ExcelDocumentService.IsValidNamePart(modelName))
            {
                ShowWarning(AppText.Get("Please enter a Model that can be used in a file name (no \\/:*?\"<>| characters or trailing period)."),
                    AppText.Get("Invalid Model"));
                textBox1.Focus();
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
                    ShowWarning(AppText.Get("Please enter a valid name in {0}.", AppText.Get(i == 0 ? "Part" : "CRC")),
                        AppText.Get("Invalid file name"));
                    nameInputs[i].Focus();
                    return;
                }

                string requiredPrefix = i == 0 ? "SAA" : "0x";
                if (!part.StartsWith(requiredPrefix, StringComparison.Ordinal))
                {
                    ShowWarning(AppText.Get("{0} must start with {1}.", AppText.Get(i == 0 ? "Part" : "CRC"), requiredPrefix),
                        AppText.Get("Invalid {0}", AppText.Get(i == 0 ? "Part" : "CRC")));
                    nameInputs[i].Focus();
                    return;
                }

                nameParts[i] = part;
            }

            selectedDateCalendar.MinDate = DateTime.Today;
            if (selectedDateCalendar.SelectionStart.Date < DateTime.Today)
            {
                ShowWarning(AppText.Get("Please select today or a future date."), AppText.Get("Invalid date"));
                return;
            }

            try
            {
                string destination = documentService.GetDestination(sourceFileTextBox.Text, new[] { modelName, nameParts[0], nameParts[1] }, selectedDateCalendar.SelectionStart);
                if (File.Exists(destination) || Directory.Exists(destination))
                {
                    ShowWarning(AppText.Get("A file or folder already exists at:\n{0}\nPlease enter a different name.", destination), AppText.Get("Name already exists"));
                    return;
                }

                documentService.Generate(sourceFileTextBox.Text, destination, powerpointTextBox.Text.Trim(), nameParts[0], nameParts[1], textBox1.Text);
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
                    ShowWarning(AppText.Get("Excel document generated:\n{0}\n\nCould not open its folder.\n{1}", destination, ex.Message), AppText.Get("Document generated"));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, AppText.Get("Could not generate the Excel document.\n{0}", ex.Message),
                    AppText.Get("Generate failed"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // The Model value is read when Generate is clicked.
        }

        private void label1_Click(object sender, EventArgs e)
        {
            // No action is needed when the Model label is clicked.
        }

        private void ShowWarning(string message, string title)
        {
            MessageBox.Show(this, message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}

