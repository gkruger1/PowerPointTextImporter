using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Linq;

namespace PowerPointTextImporter
{
    public partial class Form1 : Form
    {
        private dynamic pptApp;
        private const string RegistryKey = @"Software\PowerPointTextImporter";
        private const string DontShowWarningValue = "DontShowWarning";
        private string selectedFilePath = string.Empty;

        public Form1()
        {
            InitializeComponent();
            
            // Set the form icon
            try
            {
                var icon = IconGenerator.CreateIcon();
                if (icon != null)
                {
                    this.Icon = icon;
                }
            }
            catch (Exception ex)
            {
                // Log the error but don't show it to the user
                System.Diagnostics.Debug.WriteLine($"Failed to set icon: {ex.Message}");
            }
            
            InitializeUI();
            ShowSecurityWarningIfNeeded();
        }

        private void ShowSecurityWarningIfNeeded()
        {
            // Check if user has chosen not to see the warning
            using (var key = Registry.CurrentUser.OpenSubKey(RegistryKey))
            {
                if (key != null)
                {
                    var dontShow = key.GetValue(DontShowWarningValue);
                    if (dontShow != null && (int)dontShow == 1)
                    {
                        return;
                    }
                }
            }

            using (var form = new Form())
            {
                form.Text = "Security Information";
                form.Width = 450;
                form.Height = 200;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                var message = new Label
                {
                    Text = "This application will create presentations in PowerPoint.\n\n" +
                           "You may see a security warning from Windows asking if you want to allow this. " +
                           "This is normal and required for the application to work.",
                    Location = new Point(20, 20),
                    Width = 400,
                    Height = 80
                };

                var checkbox = new CheckBox
                {
                    Text = "Don't show this warning again",
                    Location = new Point(20, 100),
                    Width = 200
                };

                var okButton = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Location = new Point(180, 130)
                };

                form.Controls.AddRange(new Control[] { message, checkbox, okButton });
                form.AcceptButton = okButton;

                if (form.ShowDialog() == DialogResult.OK && checkbox.Checked)
                {
                    // Save the user's preference
                    using (var key = Registry.CurrentUser.CreateSubKey(RegistryKey))
                    {
                        key.SetValue(DontShowWarningValue, 1);
                    }
                }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            CleanupPowerPoint();
        }

        private void CleanupPowerPoint()
        {
            if (pptApp != null)
            {
                try
                {
                    // Check if PowerPoint is still running before trying to quit
                    try
                    {
                        var isOpen = pptApp.Visible;
                        pptApp.Quit();
                    }
                    catch
                    {
                        // PowerPoint is already closed, ignore the error
                    }
                }
                finally
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(pptApp);
                    }
                    catch
                    {
                        // Ignore any COM errors during cleanup
                    }
                    pptApp = null;
                }
            }
        }

        private void InitializeUI()
        {
            this.Text = "PowerPoint Text Importer";
            this.Width = 500;
            this.Height = 200;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            var filePathLabel = new Label
            {
                AutoSize = true,
                Location = new Point(20, 60),
                Text = "No file selected"
            };

            var selectFileButton = new Button
            {
                Text = "Select Text File",
                Width = 100,
                Location = new Point(20, 20)
            };
            selectFileButton.Click += SelectFile_Click;

            var importButton = new Button
            {
                Text = "Import to PowerPoint",
                Width = 120,
                Location = new Point(140, 20)
            };
            importButton.Click += Import_Click;

            var exampleButton = new Button
            {
                Text = "Example Format",
                Width = 100,
                Location = new Point(280, 20)
            };
            exampleButton.Click += Example_Click;

            this.Controls.AddRange(new Control[] { selectFileButton, importButton, exampleButton, filePathLabel });
        }

        private void SelectFile_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select Text File";
                openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = openFileDialog.FileName;
                    var filePathLabel = Controls.OfType<Label>().First();
                    filePathLabel.Text = $"Selected file: {System.IO.Path.GetFileName(selectedFilePath)}";
                }
            }
        }

        private void Import_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFilePath))
            {
                MessageBox.Show("Please select a text file first.", "No File Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!System.IO.File.Exists(selectedFilePath))
            {
                MessageBox.Show("The selected file no longer exists.", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                string text = System.IO.File.ReadAllText(selectedFilePath);
                CreatePowerPointPresentation(text);
                MessageBox.Show("Slides created successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error importing file: {ex.Message}", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                CleanupPowerPoint();
            }
        }

        private void CreatePowerPointPresentation(string text)
        {
            try
            {
                pptApp = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
                pptApp.Visible = true;

                dynamic presentation = pptApp.Presentations.Add();
                
                // Split text into slides
                var slideTexts = System.Text.RegularExpressions.Regex.Split(text, @"(?=Slide \d+:)").Where(s => !string.IsNullOrWhiteSpace(s));

                foreach (string slideText in slideTexts)
                {
                    // Extract slide title and content
                    var titleMatch = System.Text.RegularExpressions.Regex.Match(slideText, @"Slide \d+:(.*?)(?=[\r\n]|$)");
                    var bulletPoints = System.Text.RegularExpressions.Regex.Matches(slideText, @"^-\s*(.*?)$", System.Text.RegularExpressions.RegexOptions.Multiline)
                                          .Cast<System.Text.RegularExpressions.Match>()
                                          .Select(m => m.Groups[1].Value.Trim())
                                          .ToList();

                    if (titleMatch.Success)
                    {
                        // Add new slide (index 2 is typically the Title and Content layout)
                        dynamic slide = presentation.Slides.Add(presentation.Slides.Count + 1, 2);

                        // Add title
                        var title = titleMatch.Groups[1].Value.Trim();
                        slide.Shapes.Title.TextFrame.TextRange.Text = title;

                        // Add bullet points to the content placeholder
                        var bodyShape = slide.Shapes.Item(2);
                        var textRange = bodyShape.TextFrame.TextRange;
                        // Remove the bullet point character since PowerPoint will add it automatically
                        textRange.Text = string.Join("\n", bulletPoints);
                    }
                }
            }
            catch
            {
                CleanupPowerPoint();
                throw;
            }
        }

        private void Example_Click(object sender, EventArgs e)
        {
            string example = 
@"Slide 1: Introduction
- Welcome to our presentation
- Today's agenda
- Key topics

Slide 2: Main Points
- First important point
- Second important point
- Supporting details

Note: Each slide starts with 'Slide X:' followed by the title.
Each bullet point starts with a dash (-).";

            MessageBox.Show(example, "Example Text File Format", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
