using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelToCSVConverter1.Services
{
    public class CsvProcessingService
    {
        public DataTable ReadCSVFile(string filePath)
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
            // Process all lines to fill the data
            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue; // Skip empty lines

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

                if (rowValues.Count > 0) // Only add rows with actual data
                {
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

        public void SaveAsCSV(DataTable dataTable, string outputPath)
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