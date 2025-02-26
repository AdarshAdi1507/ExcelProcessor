using System.Collections.Generic;

namespace ExcelToCSVConverter1.Models
{
    public class ProcessingResult
    {
        public int ProcessedCount { get; set; }
        public int SkippedCount { get; set; }
        public List<string> ProcessedFiles { get; set; } = new List<string>();
        public List<string> SkippedFiles { get; set; } = new List<string>();
        public string OutputFilePath { get; set; }
    }
}