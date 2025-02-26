using System;
using System.IO;
using System.Windows.Forms;
using ExcelToCSVConverter1.Models;
using ExcelToCSVConverter1.Services;
using System.Collections.Generic;
using ExcelToCSVConverter1.Models;
using ExcelToCSVConverter1.Services;

namespace ExcelToCSVConverter1.Forms
{
    public partial class MainForm : Form
    {
        private Label lblFile, lblMode, lblStatus;
        private TextBox txtFilePath;
        private Button btnBrowse, btnProcess, btnOpenFolder, btnOpenLog, btnOpenFile;
        private RadioButton rbSingleFile, rbFolder;
        private string selectedFilePath;
        private string selectedFolderPath;
        private string outputFolderPath;
        private string logFilePath;
        private string outputFilePath;
        private FileProcessingService processingService;
        private LoggingService loggingService;

        public MainForm()
        {
            InitializeComponents();
            InitializeServices();
        }

        private void InitializeServices()
        {
            // Initialize services
            loggingService = new LoggingService();
            processingService = new FileProcessingService(
                new ExcelProcessingService(),
                new CsvProcessingService(),
                loggingService
            );
        }

        private void InitializeComponents()
        {
            this.Text = "Excel to CSV Converter";
            this.Width = 550;
            this.Height = 280;
            this.StartPosition = FormStartPosition.CenterScreen;

            // Mode Selection
            lblMode = new Label { Text = "Processing Mode:", Left = 20, Top = 20, Width = 120 };
            this.Controls.Add(lblMode);

            rbSingleFile = new RadioButton { Text = "Single File", Left = 150, Top = 20, Width = 100, Checked = true };
            rbSingleFile.CheckedChanged += new EventHandler(RbMode_CheckedChanged);
            this.Controls.Add(rbSingleFile);

            rbFolder = new RadioButton { Text = "Folder", Left = 260, Top = 20, Width = 100 };
            rbFolder.CheckedChanged += new EventHandler(RbMode_CheckedChanged);
            this.Controls.Add(rbFolder);

            // Label
            lblFile = new Label { Text = "Select File:", Left = 20, Top = 50, Width = 120 };
            this.Controls.Add(lblFile);

            // TextBox (File Path)
            txtFilePath = new TextBox { Left = 20, Top = 80, Width = 400, ReadOnly = true };
            this.Controls.Add(txtFilePath);

            // Browse Button
            btnBrowse = new Button { Text = "Browse", Left = 430, Top = 78, Width = 80 };
            btnBrowse.Click += new EventHandler(BtnBrowse_Click);
            this.Controls.Add(btnBrowse);

            // Process Button (Initially Disabled)
            btnProcess = new Button { Text = "Process", Left = 20, Top = 120, Width = 120, Enabled = false };
            btnProcess.Click += new EventHandler(BtnProcess_Click);
            this.Controls.Add(btnProcess);

            // Open Folder Button (Initially Disabled)
            btnOpenFolder = new Button { Text = "Open Output Folder", Left = 150, Top = 120, Width = 150, Enabled = false };
            btnOpenFolder.Click += new EventHandler(BtnOpenFolder_Click);
            this.Controls.Add(btnOpenFolder);

            // Open Log Button (Initially Disabled)
            btnOpenLog = new Button { Text = "Open Log File", Left = 310, Top = 120, Width = 100, Enabled = false };
            btnOpenLog.Click += new EventHandler(BtnOpenLog_Click);
            this.Controls.Add(btnOpenLog);

            // Open File Button (Initially Disabled)
            btnOpenFile = new Button { Text = "Open CSV File", Left = 420, Top = 120, Width = 100, Enabled = false };
            btnOpenFile.Click += new EventHandler(BtnOpenFile_Click);
            this.Controls.Add(btnOpenFile);

            // Status Label
            lblStatus = new Label { Text = "", Left = 20, Top = 160, Width = 500, Height = 60 };
            this.Controls.Add(lblStatus);
        }

        private void RbMode_CheckedChanged(object sender, EventArgs e)
        {
            if (rbSingleFile.Checked)
            {
                lblFile.Text = "Select File:";
                txtFilePath.Text = "";
                selectedFilePath = null;
            }
            else
            {
                lblFile.Text = "Select Folder:";
                txtFilePath.Text = "";
                selectedFolderPath = null;
            }
            btnProcess.Enabled = false;
            btnOpenFolder.Enabled = false;
            btnOpenLog.Enabled = false;
            btnOpenFile.Enabled = false;
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            if (rbSingleFile.Checked)
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx, *.csv)|*.xlsx;*.csv",
                    Title = "Select an Excel or CSV File"
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = openFileDialog.FileName;
                    txtFilePath.Text = selectedFilePath;
                    btnProcess.Enabled = true;
                }
            }
            else
            {
                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog
                {
                    Description = "Select a folder containing Excel or CSV files"
                };

                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFolderPath = folderBrowserDialog.SelectedPath;
                    txtFilePath.Text = selectedFolderPath;
                    btnProcess.Enabled = true;
                }
            }
        }

        private void BtnProcess_Click(object sender, EventArgs e)
        {
            try
            {
                btnProcess.Enabled = false;
                btnOpenFolder.Enabled = false;
                btnOpenLog.Enabled = false;
                btnOpenFile.Enabled = false;
                lblStatus.Text = "Processing...";
                Application.DoEvents();

                ProcessingOptions options = new ProcessingOptions
                {
                    PriorityColumns = new List<string>
                    {
                    "sName",
                    "sRevision",
                    "sType",
                    "CAD Type",
                    "Originator",
                    "Title",
                    "MCADInteg-Comment",
                    "HSIDRWPARTNAME",
                    "sDescription",
                    "sfileNames",
                    "sfileFormats",
                    "snooffileFormats",
                    "sfilefolderPath"
                    }
                };

                ProcessingResult result;

                if (rbSingleFile.Checked)
                {
                    // Process single file
                    string directory = Path.GetDirectoryName(selectedFilePath);
                    outputFolderPath = Path.Combine(directory, "output");

                    if (!Directory.Exists(outputFolderPath))
                    {
                        Directory.CreateDirectory(outputFolderPath);
                    }

                    logFilePath = Path.Combine(outputFolderPath, "conversion_log.txt");
                    loggingService.Initialize(logFilePath);

                    result = processingService.ProcessSingleFile(selectedFilePath, outputFolderPath, options);

                    // Set output file path for single file mode
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(selectedFilePath);
                    outputFilePath = Path.Combine(outputFolderPath, fileNameWithoutExt + "_converted.csv");
                }
                else
                {
                    // Process all files in folder
                    outputFolderPath = Path.Combine(selectedFolderPath, "output");

                    if (!Directory.Exists(outputFolderPath))
                    {
                        Directory.CreateDirectory(outputFolderPath);
                    }

                    logFilePath = Path.Combine(outputFolderPath, "conversion_log.txt");
                    loggingService.Initialize(logFilePath);

                    result = processingService.ProcessFolder(selectedFolderPath, outputFolderPath, options);

                    // Get the output file path from the result
                    outputFilePath = result.OutputFilePath;
                }

                // Update UI
                lblStatus.Text = $"Processing completed!\n" +
                                 $"Files processed: {result.ProcessedCount}\n" +
                                 $"Files skipped: {result.SkippedCount}\n" +
                                 $"Check log file for details.";

                btnProcess.Enabled = true;
                btnOpenFolder.Enabled = true;
                btnOpenLog.Enabled = true;
                btnOpenFile.Enabled = outputFilePath != null && File.Exists(outputFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Processing Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnProcess.Enabled = true;
            }
        }

        private void BtnOpenFolder_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(outputFolderPath))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = outputFolderPath,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Output folder not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnOpenLog_Click(object sender, EventArgs e)
        {
            if (File.Exists(logFilePath))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = logFilePath,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Log file not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnOpenFile_Click(object sender, EventArgs e)
        {
            if (File.Exists(outputFilePath))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = outputFilePath,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("CSV file not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}