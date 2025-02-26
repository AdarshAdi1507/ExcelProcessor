using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelToCSVConverter1.Services
{
    public class ExcelProcessingService
    {
        public DataTable ReadExcelFile(string filePath)
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
                // Second pass: fill the data
                foreach (var row in worksheet.RowsUsed())
                {
                    DataRow dataRow = dataTable.NewRow();
                    Dictionary<string, string> rowValues = new Dictionary<string, string>();

                    string fullRowText = string.Join("\t", row.CellsUsed().Select(cell => cell.Value.ToString())).Trim();
                    if (string.IsNullOrWhiteSpace(fullRowText)) continue; // Skip empty rows

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
            }

                return dataTable;
        }
    }
}































