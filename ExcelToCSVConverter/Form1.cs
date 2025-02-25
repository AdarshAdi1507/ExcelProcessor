using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Data;
using System.Diagnostics;
using System.Collections.Generic;

namespace ExcelToCSVConverter
{
    public partial class Form1 : Form
    {
        private Label lblFile;
        private TextBox txtFilePath;
        private Button btnBrowse, btnConvert, btnOpenFile;
        private string selectedFilePath;
        private string outputFilePath;

        // Define priority columns in the desired order (these will appear first)
        private List<string> priorityColumns = new List<string>
        {
            "sName",
            "sRevision",
            "sType",
            "CAD Type",
            "Originator",
            "Title" 
            ,"MCADInteg-Comment",
            "HSIDRWPARTNAME"
            // You can easily modify this list to change the order or add/remove columns
        };

        public Form1()
        {
            this.Text = "Excel to CSV Converter";  // Set window title
            this.Width = 500;
            this.Height = 200;
            this.StartPosition = FormStartPosition.CenterScreen;

            // Label
            lblFile = new Label { Text = "Select an Excel/CSV File:", Left = 20, Top = 20, Width = 200 };
            this.Controls.Add(lblFile);

            // TextBox (File Path)
            txtFilePath = new TextBox { Left = 20, Top = 50, Width = 350, ReadOnly = true };
            this.Controls.Add(txtFilePath);

            // Browse Button
            btnBrowse = new Button { Text = "Browse", Left = 380, Top = 47, Width = 80 };
            btnBrowse.Click += new EventHandler(BtnBrowse_Click);
            this.Controls.Add(btnBrowse);

            // Convert Button (Initially Disabled)
            btnConvert = new Button { Text = "Convert", Left = 20, Top = 90, Width = 100, Enabled = false };
            btnConvert.Click += new EventHandler(BtnConvert_Click);
            this.Controls.Add(btnConvert);

            // Open File Button (Initially Disabled)
            btnOpenFile = new Button { Text = "Open File", Left = 140, Top = 90, Width = 100, Enabled = false };
            btnOpenFile.Click += new EventHandler(BtnOpenFile_Click);
            this.Controls.Add(btnOpenFile);
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
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
                btnConvert.Enabled = true;  // Enable Convert button
            }
        }

        private void BtnConvert_Click(object sender, EventArgs e)
        {
            try
            {
                // Generate output file name
                string directory = Path.GetDirectoryName(selectedFilePath);
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(selectedFilePath);
                outputFilePath = Path.Combine(directory, fileNameWithoutExt + "_converted.csv");

                string extension = Path.GetExtension(selectedFilePath).ToLower();
                DataTable dataTable;

                if (extension == ".csv")
                {
                    dataTable = ReadCSVFile(selectedFilePath);
                }
                else
                {
                    dataTable = ReadExcelFile(selectedFilePath);
                }

                // Reorder the columns before saving
                dataTable = ReorderColumns(dataTable);

                SaveAsCSV(dataTable, outputFilePath);

                MessageBox.Show("Conversion Successful! File saved as: " + outputFilePath, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnOpenFile.Enabled = true; // Enable Open File button
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Conversion Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnOpenFile_Click(object sender, EventArgs e)
        {
            if (File.Exists(outputFilePath))
            {
                Process.Start(new ProcessStartInfo(outputFilePath) { UseShellExecute = true });
            }
            else
            {
                MessageBox.Show("File not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private DataTable ReadCSVFile(string filePath)
        {
            DataTable dataTable = new DataTable();
            List<string> headers = new List<string>();

            string[] lines = File.ReadAllLines(filePath);

            if (lines.Length == 0)
                return dataTable;

            // Process first line to detect all possible headers
            foreach (var item in lines[0].Split('\t'))
            {
                string[] keyValue = item.Split(new[] { '=' }, 2);
                if (keyValue.Length == 2)
                {
                    string header = keyValue[0].Trim();
                    if (!headers.Contains(header))
                    {
                        headers.Add(header);
                        dataTable.Columns.Add(header);
                    }
                }
            }

            // Process all lines to find any additional headers
            foreach (var line in lines)
            {
                foreach (var item in line.Split('\t'))
                {
                    string[] keyValue = item.Split(new[] { '=' }, 2);
                    if (keyValue.Length == 2)
                    {
                        string header = keyValue[0].Trim();
                        if (!headers.Contains(header))
                        {
                            headers.Add(header);
                            dataTable.Columns.Add(header);
                        }
                    }
                }
            }

            // Process all lines to fill the data
            foreach (var line in lines)
            {
                DataRow dataRow = dataTable.NewRow();
                Dictionary<string, string> rowValues = new Dictionary<string, string>();

                foreach (var item in line.Split('\t'))
                {
                    string[] keyValue = item.Split(new[] { '=' }, 2);
                    if (keyValue.Length == 2)
                    {
                        string header = keyValue[0].Trim();
                        string value = keyValue[1].Trim();
                        rowValues[header] = value;
                    }
                }

                foreach (string header in headers)
                {
                    if (rowValues.ContainsKey(header))
                    {
                        dataRow[header] = rowValues[header];
                    }
                }

                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }

        private DataTable ReadExcelFile(string filePath)
        {
            DataTable dataTable = new DataTable();
            List<string> headers = new List<string>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.First();

                // First pass: collect all possible headers
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        string cellValue = cell.Value.ToString();
                        string[] parts = cellValue.Split(new[] { '\t' });

                        foreach (var part in parts)
                        {
                            string[] keyValue = part.Split(new[] { '=' }, 2);
                            if (keyValue.Length == 2)
                            {
                                string header = keyValue[0].Trim();
                                if (!headers.Contains(header))
                                {
                                    headers.Add(header);
                                    dataTable.Columns.Add(header);
                                }
                            }
                        }
                    }
                }

                // Second pass: fill the data
                foreach (var row in worksheet.RowsUsed())
                {
                    DataRow dataRow = dataTable.NewRow();
                    Dictionary<string, string> rowValues = new Dictionary<string, string>();

                    string fullRowText = "";
                    foreach (var cell in row.CellsUsed())
                    {
                        fullRowText += cell.Value.ToString() + "\t";
                    }

                    string[] parts = fullRowText.Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (var part in parts)
                    {
                        string[] keyValue = part.Split(new[] { '=' }, 2);
                        if (keyValue.Length == 2)
                        {
                            string header = keyValue[0].Trim();
                            string value = keyValue[1].Trim();
                            rowValues[header] = value;
                        }
                    }

                    foreach (string header in headers)
                    {
                        if (rowValues.ContainsKey(header))
                        {
                            dataRow[header] = rowValues[header];
                        }
                    }

                    dataTable.Rows.Add(dataRow);
                }
            }

            return dataTable;
        }

        private DataTable ReorderColumns(DataTable originalTable)
        {
            // Create a new DataTable to hold the reordered data
            DataTable reorderedTable = new DataTable();

            // First add the priority columns in the specified order (if they exist in the original table)
            foreach (string colName in priorityColumns)
            {
                if (originalTable.Columns.Contains(colName))
                {
                    // Add the column to the new table
                    reorderedTable.Columns.Add(colName, originalTable.Columns[colName].DataType);
                }
            }

            // Then add all remaining columns that weren't in the priority list
            foreach (DataColumn col in originalTable.Columns)
            {
                if (!priorityColumns.Contains(col.ColumnName))
                {
                    reorderedTable.Columns.Add(col.ColumnName, col.DataType);
                }
            }

            // Copy the data from the original table to the reordered table
            foreach (DataRow originalRow in originalTable.Rows)
            {
                DataRow newRow = reorderedTable.NewRow();

                foreach (DataColumn col in reorderedTable.Columns)
                {
                    newRow[col.ColumnName] = originalRow[col.ColumnName];
                }

                reorderedTable.Rows.Add(newRow);
            }

            return reorderedTable;
        }

        private void SaveAsCSV(DataTable dataTable, string outputPath)
        {
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                // Write headers
                writer.WriteLine(string.Join(",", dataTable.Columns.Cast<DataColumn>().Select(c => EscapeCSVField(c.ColumnName))));

                // Write rows
                foreach (DataRow row in dataTable.Rows)
                {
                    writer.WriteLine(string.Join(",", row.ItemArray.Select(field => EscapeCSVField(field.ToString()))));
                }
            }
        }

        private string EscapeCSVField(string field)
        {
            // If the field contains comma, newline, or double quote, escape it
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                // Replace any double quotes with two double quotes
                field = field.Replace("\"", "\"\"");
                // Enclose the field in double quotes
                field = "\"" + field + "\"";
            }
            return field;
        }
    }
}