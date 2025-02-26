using System.Collections.Generic;

namespace ExcelToCSVConverter1.Models
{
    public class ProcessingOptions
    {
        // Define priority columns in the desired order (these will appear first)
        public List<string> PriorityColumns { get; set; } = new List<string>();
    }
}