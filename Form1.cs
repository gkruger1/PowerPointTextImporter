using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Linq;
using System.Collections.Generic;

namespace PowerPointTextImporter
{
    public partial class Form1 : Form
    {
        private dynamic pptApp;
        private bool wasAlreadyOpen = false;
        private dynamic currentPresentation = null;
        private bool userExplicitlyDeclinedSave = false;
        private bool userCanceled = false;
        private const string RegistryKey = @"Software\PowerPointTextImporter";
        private const string DontShowWarningValue = "DontShowWarning";
        private string selectedFilePath = string.Empty;
        private FlowLayoutPanel previewPanel;
        private List<CheckBox> slideCheckboxes;
        private ToolTip tooltip;

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
            
            // Initialize the tooltip
            tooltip = new ToolTip
            {
                InitialDelay = 0,
                ReshowDelay = 0,
                AutoPopDelay = 32000,  // Keep tooltip visible for longer
                ShowAlways = true,
                IsBalloon = true,
                UseFading = false,  // Disable fading effect
                UseAnimation = false  // Disable animation
            };
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
            if (currentPresentation != null)
            {
                try
                {
                    if (!wasAlreadyOpen && !userExplicitlyDeclinedSave)
                    {
                        // If presentation wasn't saved and we created it, ask one more time
                        if (currentPresentation.Saved == false)
                        {
                            var saveResult = MessageBox.Show(
                                "Do you want to save changes to the presentation?",
                                "Save Changes",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Question
                            );

                            if (saveResult == DialogResult.Yes)
                            {
                                using (var saveDialog = new SaveFileDialog())
                                {
                                    saveDialog.Filter = "PowerPoint Presentation (*.pptx)|*.pptx";
                                    saveDialog.DefaultExt = "pptx";
                                    saveDialog.AddExtension = true;

                                    if (saveDialog.ShowDialog() == DialogResult.OK)
                                    {
                                        currentPresentation.SaveAs(saveDialog.FileName);
                                    }
                                    else
                                    {
                                        userCanceled = true; // User canceled the save dialog
                                    }
                                }
                            }
                            else if (saveResult == DialogResult.Cancel)
                            {
                                userCanceled = true;
                            }
                        }
                    }

                    // Only release the presentation if we're not keeping PowerPoint open
                    if (!userCanceled)
                    {
                        Marshal.FinalReleaseComObject(currentPresentation);
                        currentPresentation = null;
                    }
                }
                catch
                {
                    // Ignore any COM errors during cleanup
                }
            }

            if (pptApp != null && !userCanceled) // Only cleanup PowerPoint if not canceled
            {
                try
                {
                    if (!wasAlreadyOpen)
                    {
                        try
                        {
                            pptApp.Quit();
                        }
                        catch
                        {
                            // PowerPoint is already closed, ignore the error
                        }
                    }

                    Marshal.FinalReleaseComObject(pptApp);
                    pptApp = null;
                }
                catch
                {
                    // Ignore any COM errors during cleanup
                }
            }
        }

        private void InitializeUI()
        {
            this.Text = "PowerPoint Text Importer";
            this.Width = 500;
            this.Height = 400;  // Increased height for preview
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

            // Add preview panel
            previewPanel = new FlowLayoutPanel
            {
                Location = new Point(20, 90),
                Width = 440,
                Height = 250,
                AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Visible = false
            };

            this.Controls.AddRange(new Control[] { selectFileButton, importButton, exampleButton, filePathLabel, previewPanel });
            slideCheckboxes = new List<CheckBox>();
        }

        private class SlideValidationResult
        {
            public bool IsValid { get; set; }
            public string Message { get; set; }
            public string Title { get; set; }
            public List<string> BulletPoints { get; set; }
        }

        private SlideValidationResult ValidateSlide(string slideText)
        {
            var result = new SlideValidationResult
            {
                IsValid = true,
                BulletPoints = new List<string>()
            };

            // Validate title
            var titleMatch = Regex.Match(slideText, @"Slide \d+:(.*?)(?=[\r\n]|$)");
            if (!titleMatch.Success || string.IsNullOrWhiteSpace(titleMatch.Groups[1].Value))
            {
                result.IsValid = false;
                result.Message = "Invalid slide title format";
                result.Title = slideText.Split('\n')[0];
                return result;
            }
            result.Title = titleMatch.Value.Trim();

            // Extract and validate bullet points
            var bulletPoints = Regex.Matches(slideText, @"^-\s*(.*?)$", RegexOptions.Multiline)
                                  .Cast<Match>()
                                  .Select(m => m.Groups[1].Value.Trim())
                                  .Where(bp => !string.IsNullOrWhiteSpace(bp))
                                  .ToList();

            if (bulletPoints.Count == 0)
            {
                result.IsValid = false;
                result.Message = "No bullet points found";
                return result;
            }

            result.BulletPoints = bulletPoints;
            return result;
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
                    UpdatePreview();
                }
            }
        }

        private void UpdatePreview()
        {
            try
            {
                string text = System.IO.File.ReadAllText(selectedFilePath);
                
                // First split by double newlines to separate potential slides
                var slideTexts = text.Split(new[] { "\r\n\r\n", "\n\n" }, StringSplitOptions.RemoveEmptyEntries)
                                   .Where(s => !string.IsNullOrWhiteSpace(s))
                                   .ToList();

                previewPanel.Controls.Clear();
                slideCheckboxes.Clear();

                var selectAllCheckbox = new CheckBox
                {
                    Text = "Select All Slides",
                    Checked = true,
                    AutoSize = true,
                    Margin = new Padding(5)
                };
                selectAllCheckbox.CheckedChanged += (s, e) =>
                {
                    foreach (var cb in slideCheckboxes)
                    {
                        if (cb.Enabled) // Only change enabled checkboxes
                            cb.Checked = selectAllCheckbox.Checked;
                    }
                };
                previewPanel.Controls.Add(selectAllCheckbox);

                foreach (string slideText in slideTexts)
                {
                    var validationResult = ValidateSlide(slideText);
                    
                    // Create panel for the slide preview row
                    var slidePanel = new Panel
                    {
                        Width = previewPanel.Width - 20,
                        Height = 30,
                        Margin = new Padding(5),
                        Cursor = validationResult.IsValid ? Cursors.Default : Cursors.Help,
                        BackColor = Color.Transparent
                    };

                    // Add validation indicator
                    var indicator = new Label
                    {
                        AutoSize = true,
                        Location = new Point(0, 5),
                        Text = validationResult.IsValid ? "✓" : "✗",
                        ForeColor = validationResult.IsValid ? Color.Green : Color.Red,
                        Font = new Font(Font.FontFamily, 10, FontStyle.Bold),
                        Cursor = validationResult.IsValid ? Cursors.Default : Cursors.Help
                    };
                    slidePanel.Controls.Add(indicator);

                    if (validationResult.IsValid)
                    {
                        // Add checkbox for valid slides
                        var checkbox = new CheckBox
                        {
                            Text = validationResult.Title,
                            Checked = true,
                            AutoSize = true,
                            Location = new Point(20, 5),
                            Enabled = true,
                            Cursor = Cursors.Default
                        };
                        slideCheckboxes.Add(checkbox);
                        slidePanel.Controls.Add(checkbox);

                        string tooltipText = $"Valid slide with {validationResult.BulletPoints.Count} bullet points:\n\n{string.Join("\n", validationResult.BulletPoints)}";
                        tooltip.SetToolTip(indicator, tooltipText);
                        tooltip.SetToolTip(checkbox, tooltipText);
                    }
                    else
                    {
                        // For invalid slides, use a label with checkbox appearance
                        var checkboxLabel = new Label
                        {
                            Text = "☐ Invalid Format: " + slideText.Split('\n')[0].Trim(),
                            AutoSize = true,
                            Location = new Point(20, 5),
                            Cursor = Cursors.Help
                        };
                        slidePanel.Controls.Add(checkboxLabel);
                        
                        // Add a disabled checkbox for tracking (invisible)
                        var hiddenCheckbox = new CheckBox
                        {
                            Visible = false,
                            Checked = false,
                            Enabled = false
                        };
                        slideCheckboxes.Add(hiddenCheckbox);
                        slidePanel.Controls.Add(hiddenCheckbox);

                        string tooltipText = 
$@"Error: {validationResult.Message}

Expected format:
Slide 1: Your Title
- Bullet point 1
- Bullet point 2

Your text:
{slideText.Trim()}";

                        tooltip.SetToolTip(indicator, tooltipText);
                        tooltip.SetToolTip(checkboxLabel, tooltipText);
                        tooltip.SetToolTip(slidePanel, tooltipText);
                    }

                    previewPanel.Controls.Add(slidePanel);
                }

                previewPanel.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error previewing file: {ex.Message}", "Preview Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                // Check if PowerPoint is already running
                try
                {
                    pptApp = System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application");
                    wasAlreadyOpen = true;
                }
                catch
                {
                    pptApp = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
                    wasAlreadyOpen = false;
                }

                pptApp.Visible = true;
                currentPresentation = pptApp.Presentations.Add();
                
                // Split text into slides
                var slideTexts = text.Split(new[] { "\r\n\r\n", "\n\n" }, StringSplitOptions.RemoveEmptyEntries)
                                   .Where(s => !string.IsNullOrWhiteSpace(s))
                                   .ToList();

                for (int i = 0; i < slideTexts.Count; i++)
                {
                    // Skip if the corresponding checkbox is unchecked
                    if (i < slideCheckboxes.Count && !slideCheckboxes[i].Checked)
                        continue;

                    string slideText = slideTexts[i];
                    var validationResult = ValidateSlide(slideText);
                    
                    // Skip invalid slides
                    if (!validationResult.IsValid)
                        continue;

                    // Add new slide (index 2 is typically the Title and Content layout)
                    dynamic slide = currentPresentation.Slides.Add(currentPresentation.Slides.Count + 1, 2);

                    // Add title
                    var title = validationResult.Title.Replace("Slide " + (i + 1) + ":", "").Trim();
                    slide.Shapes.Title.TextFrame.TextRange.Text = title;

                    // Add bullet points to the content placeholder
                    var bodyShape = slide.Shapes.Item(2);
                    var textRange = bodyShape.TextFrame.TextRange;
                    textRange.Text = string.Join("\n", validationResult.BulletPoints);
                }

                // Ask user if they want to save the presentation
                var saveResult = MessageBox.Show(
                    "Would you like to save the presentation?",
                    "Save Presentation",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question
                );

                if (saveResult == DialogResult.Yes)
                {
                    using (var saveDialog = new SaveFileDialog())
                    {
                        saveDialog.Filter = "PowerPoint Presentation (*.pptx)|*.pptx";
                        saveDialog.DefaultExt = "pptx";
                        saveDialog.AddExtension = true;

                        if (saveDialog.ShowDialog() == DialogResult.OK)
                        {
                            currentPresentation.SaveAs(saveDialog.FileName);
                            MessageBox.Show("Presentation saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            userExplicitlyDeclinedSave = true; // Mark as handled since we saved successfully
                        }
                        else
                        {
                            userCanceled = true; // User canceled the save dialog
                        }
                    }
                }
                else if (saveResult == DialogResult.No)
                {
                    userExplicitlyDeclinedSave = true; // User explicitly chose not to save
                }
                else // Cancel
                {
                    userCanceled = true;
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
