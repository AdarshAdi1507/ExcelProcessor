using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ExcelToCSVConverter1.Models;
using ExcelToCSVConverter1.Models;
using ExcelToCSVConverter1.Services;

namespace ExcelToCSVConverter1.Services
{
    public class FileProcessingService
    {
        private readonly ExcelProcessingService _excelService;
        private readonly CsvProcessingService _csvService;
        private readonly LoggingService _loggingService;

        public FileProcessingService(
            ExcelProcessingService excelService,
            CsvProcessingService csvService,
            LoggingService loggingService)
        {
            _excelService = excelService;
            _csvService = csvService;
            _loggingService = loggingService;
        }

        public ProcessingResult ProcessSingleFile(string filePath, string outputFolderPath, ProcessingOptions options)
        {
            ProcessingResult result = new ProcessingResult();
            _loggingService.Log($"Starting to process file: {filePath}");

            try
            {
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                string outputFilePath = Path.Combine(outputFolderPath, fileNameWithoutExt + "_converted.csv");
                string extension = Path.GetExtension(filePath).ToLower();

                DataTable dataTable;
                if (extension == ".csv")
                {
                    dataTable = _csvService.ReadCSVFile(filePath);
                }
                else
                {
                    dataTable = _excelService.ReadExcelFile(filePath);
                }

                // Reorder the columns
                dataTable = ReorderColumns(dataTable, options.PriorityColumns);

                // Save as CSV
                _csvService.SaveAsCSV(dataTable, outputFilePath);

                result.ProcessedCount = 1;
                result.ProcessedFiles.Add(filePath);
                _loggingService.Log($"Successfully processed file: {filePath}");
                _loggingService.Log($"Output saved to: {outputFilePath}");
            }
            catch (Exception ex)
            {
                result.SkippedCount = 1;
                result.SkippedFiles.Add(filePath);
                _loggingService.Log($"Error processing file {filePath}: {ex.Message}");
            }

            return result;
        }

        public ProcessingResult ProcessFolder(string folderPath, string outputFolderPath, ProcessingOptions options)
        {
            ProcessingResult result = new ProcessingResult();
            _loggingService.Log($"Starting to process folder: {folderPath}");

            // Get all Excel and CSV files
            string[] filePaths = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                .Where(file => file.ToLower().EndsWith(".xlsx") || file.ToLower().EndsWith(".csv"))
                .ToArray();

            // Sort files alphanumerically
            Array.Sort(filePaths, new AlphanumericComparer());

            _loggingService.Log($"Found {filePaths.Length} files to process in alphanumeric order");

            // Create a merged data table
            DataTable mergedDataTable = null;
            int totalRowsProcessed = 0;

            foreach (string filePath in filePaths)
            {
                try
                {
                    _loggingService.Log($"Processing file: {filePath}");

                    string extension = Path.GetExtension(filePath).ToLower();
                    DataTable currentDataTable;

                    if (extension == ".csv")
                    {
                        currentDataTable = _csvService.ReadCSVFile(filePath);
                    }
                    else
                    {
                        currentDataTable = _excelService.ReadExcelFile(filePath);
                    }

                    // If this is the first file, use it to initialize the merged table structure
                    if (mergedDataTable == null)
                    {
                        // Initialize merged table with the first file's schema and priority columns
                        mergedDataTable = ReorderColumns(currentDataTable.Clone(), options.PriorityColumns);
                        _loggingService.Log($"Created merged table structure with {mergedDataTable.Columns.Count} columns");
                    }

                    foreach (DataRow row in currentDataTable.Rows)
                    {
                        // Check if the row has actual data
                        if (!row.ItemArray.All(field => string.IsNullOrWhiteSpace(field?.ToString())))
                        {
                            DataRow newRow = mergedDataTable.NewRow();

                            // Copy values for columns that exist in both tables
                            foreach (DataColumn col in mergedDataTable.Columns)
                            {
                                if (currentDataTable.Columns.Contains(col.ColumnName))
                                {
                                    newRow[col.ColumnName] = row[col.ColumnName];
                                }
                            }

                            mergedDataTable.Rows.Add(newRow);
                        }
                    }
                        totalRowsProcessed += currentDataTable.Rows.Count;
                    _loggingService.Log($"Added {currentDataTable.Rows.Count} rows from file, total rows now: {totalRowsProcessed}");

                    result.ProcessedCount++;
                    result.ProcessedFiles.Add(filePath);
                    _loggingService.Log($"Successfully processed file: {filePath}");
                }
                catch (Exception ex)
                {
                    result.SkippedCount++;
                    result.SkippedFiles.Add(filePath);
                    _loggingService.Log($"Error processing file {filePath}: {ex.Message}");
                }
            }

            // Save the merged data
            if (mergedDataTable != null && mergedDataTable.Rows.Count > 0)
            {
                string mergedOutputPath = Path.Combine(outputFolderPath, "merged_output.csv");
                _csvService.SaveAsCSV(mergedDataTable, mergedOutputPath);
                _loggingService.Log($"Saved merged output with {mergedDataTable.Rows.Count} rows to: {mergedOutputPath}");

                // Set the result's output file for the form to use
                result.OutputFilePath = mergedOutputPath;
            }
            else
            {
                _loggingService.Log("No data was processed or merged.");
            }

            _loggingService.Log($"Folder processing completed. Processed: {result.ProcessedCount}, Skipped: {result.SkippedCount}");
            return result;
        }

        private DataTable ReorderColumns(DataTable originalTable, List<string> priorityColumns)
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
    }

    // Helper class for alphanumeric sorting of filenames
    public class AlphanumericComparer : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return -1;
            if (y == null) return 1;

            int len1 = x.Length;
            int len2 = y.Length;
            int marker1 = 0;
            int marker2 = 0;

            // Walk through two strings with two markers
            while (marker1 < len1 && marker2 < len2)
            {
                char ch1 = x[marker1];
                char ch2 = y[marker2];

                // Get two chunks of the two strings
                char[] space1 = new char[len1];
                int spacePos1 = 0;
                char[] space2 = new char[len2];
                int spacePos2 = 0;

                // Walk through all digits or non-digits
                do
                {
                    space1[spacePos1++] = ch1;
                    marker1++;

                    if (marker1 < len1)
                        ch1 = x[marker1];
                    else
                        break;
                } while (char.IsDigit(ch1) == char.IsDigit(space1[0]));

                do
                {
                    space2[spacePos2++] = ch2;
                    marker2++;

                    if (marker2 < len2)
                        ch2 = y[marker2];
                    else
                        break;
                } while (char.IsDigit(ch2) == char.IsDigit(space2[0]));

                // If we have digits, compare them numerically
                string chunk1 = new string(space1, 0, spacePos1);
                string chunk2 = new string(space2, 0, spacePos2);

                int result;
                if (char.IsDigit(space1[0]) && char.IsDigit(space2[0]))
                {
                    // Extract contiguous digits
                    int numChunk1, numChunk2;
                    if (int.TryParse(chunk1, out numChunk1) && int.TryParse(chunk2, out numChunk2))
                    {
                        result = numChunk1.CompareTo(numChunk2);
                        if (result != 0)
                            return result;
                    }
                }
                else
                {
                    result = string.Compare(chunk1, chunk2, StringComparison.OrdinalIgnoreCase);
                    if (result != 0)
                        return result;
                }
            }

            return len1 - len2;
        }
    }
}