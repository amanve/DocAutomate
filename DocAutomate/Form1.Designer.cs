namespace DocAutomate
{
    partial class Form1
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
            this.browseButton = new System.Windows.Forms.Button();
            this.sourceFileTextBox = new System.Windows.Forms.TextBox();
            this.firstNameTextBox = new System.Windows.Forms.TextBox();
            this.generateButton = new System.Windows.Forms.Button();
            this.secondNameTextBox = new System.Windows.Forms.TextBox();
            this.sourceFileLabel = new System.Windows.Forms.Label();
            this.firstNameLabel = new System.Windows.Forms.Label();
            this.secondNameLabel = new System.Windows.Forms.Label();
            this.selectedDateCalendar = new System.Windows.Forms.MonthCalendar();
            this.powerpointTextBox = new System.Windows.Forms.TextBox();
            this.powerpointBrowseButton = new System.Windows.Forms.Button();
            this.powerpointLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            this.powerpointLabel.AutoSize = true;
            this.powerpointLabel.Location = new System.Drawing.Point(90, 148);
            this.powerpointLabel.Text = "PowerPoint file";
            this.powerpointTextBox.Location = new System.Drawing.Point(223, 145);
            this.powerpointTextBox.Size = new System.Drawing.Size(293, 22);
            this.powerpointTextBox.TabIndex = 6;
            this.powerpointBrowseButton.Location = new System.Drawing.Point(576, 145);
            this.powerpointBrowseButton.Size = new System.Drawing.Size(75, 23);
            this.powerpointBrowseButton.Text = "Browse";
            this.powerpointBrowseButton.TabIndex = 7;
            this.powerpointBrowseButton.Click += new System.EventHandler(this.powerpointBrowseButton_Click);
            this.Controls.Add(this.powerpointLabel);
            this.Controls.Add(this.powerpointTextBox);
            this.Controls.Add(this.powerpointBrowseButton);
            // 
            // browseButton
            // 
            this.browseButton.Location = new System.Drawing.Point(576, 102);
            this.browseButton.Name = "browseButton";
            this.browseButton.Size = new System.Drawing.Size(75, 23);
            this.browseButton.TabIndex = 0;
            this.browseButton.Text = "Browse";
            this.browseButton.UseVisualStyleBackColor = true;
            this.browseButton.Click += new System.EventHandler(this.browseButton_Click);
            // 
            // sourceFileTextBox
            // 
            this.sourceFileTextBox.Location = new System.Drawing.Point(223, 102);
            this.sourceFileTextBox.Name = "sourceFileTextBox";
            this.sourceFileTextBox.ReadOnly = true;
            this.sourceFileTextBox.Size = new System.Drawing.Size(293, 22);
            this.sourceFileTextBox.TabIndex = 1;
            // 
            // firstNameTextBox
            // 
            this.firstNameTextBox.Location = new System.Drawing.Point(223, 188);
            this.firstNameTextBox.Name = "firstNameTextBox";
            this.firstNameTextBox.Size = new System.Drawing.Size(293, 22);
            this.firstNameTextBox.TabIndex = 2;
            // 
            // generateButton
            // 
            this.generateButton.Location = new System.Drawing.Point(576, 338);
            this.generateButton.Name = "generateButton";
            this.generateButton.Size = new System.Drawing.Size(95, 34);
            this.generateButton.TabIndex = 3;
            this.generateButton.Text = "Generate";
            this.generateButton.UseVisualStyleBackColor = true;
            this.generateButton.Click += new System.EventHandler(this.generateButton_Click);
            // 
            // secondNameTextBox
            // 
            this.secondNameTextBox.Location = new System.Drawing.Point(223, 260);
            this.secondNameTextBox.Name = "secondNameTextBox";
            this.secondNameTextBox.Size = new System.Drawing.Size(293, 22);
            this.secondNameTextBox.TabIndex = 4;
            // 
            // sourceFileLabel
            // 
            this.sourceFileLabel.AutoSize = true;
            this.sourceFileLabel.Location = new System.Drawing.Point(90, 105);
            this.sourceFileLabel.Name = "sourceFileLabel";
            this.sourceFileLabel.Size = new System.Drawing.Size(63, 17);
            this.sourceFileLabel.TabIndex = 0;
            this.sourceFileLabel.Text = "Excel file";
            // 
            // firstNameLabel
            // 
            this.firstNameLabel.AutoSize = true;
            this.firstNameLabel.Location = new System.Drawing.Point(90, 191);
            this.firstNameLabel.Name = "firstNameLabel";
            this.firstNameLabel.Size = new System.Drawing.Size(57, 17);
            this.firstNameLabel.TabIndex = 1;
            this.firstNameLabel.Text = "Name 1";
            // 
            // secondNameLabel
            // 
            this.secondNameLabel.AutoSize = true;
            this.secondNameLabel.Location = new System.Drawing.Point(90, 263);
            this.secondNameLabel.Name = "secondNameLabel";
            this.secondNameLabel.Size = new System.Drawing.Size(57, 17);
            this.secondNameLabel.TabIndex = 2;
            this.secondNameLabel.Text = "Name 2";
            // 
            // selectedDateCalendar
            // 
            this.selectedDateCalendar.Location = new System.Drawing.Point(93, 294);
            this.selectedDateCalendar.Name = "selectedDateCalendar";
            this.selectedDateCalendar.TabIndex = 5;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(825, 517);
            this.Controls.Add(this.selectedDateCalendar);
            this.Controls.Add(this.sourceFileLabel);
            this.Controls.Add(this.firstNameLabel);
            this.Controls.Add(this.secondNameLabel);
            this.Controls.Add(this.secondNameTextBox);
            this.Controls.Add(this.generateButton);
            this.Controls.Add(this.firstNameTextBox);
            this.Controls.Add(this.sourceFileTextBox);
            this.Controls.Add(this.browseButton);
            this.Name = "Form1";
            this.Text = "Excel Document Generator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox powerpointTextBox;
        private System.Windows.Forms.Button powerpointBrowseButton;
        private System.Windows.Forms.Label powerpointLabel;
        private System.Windows.Forms.Label sourceFileLabel;
        private System.Windows.Forms.Label firstNameLabel;
        private System.Windows.Forms.Label secondNameLabel;
        private System.Windows.Forms.Button browseButton;
        private System.Windows.Forms.TextBox sourceFileTextBox;
        private System.Windows.Forms.TextBox firstNameTextBox;
        private System.Windows.Forms.Button generateButton;
        private System.Windows.Forms.TextBox secondNameTextBox;
        private System.Windows.Forms.MonthCalendar selectedDateCalendar;
    }
}




