using System;
using System.IO;

namespace ExcelToCSVConverter1.Services
{
    public class LoggingService
    {
        private string _logFilePath;

        public void Initialize(string logFilePath)
        {
            _logFilePath = logFilePath;

            // Create or clear the log file
            using (StreamWriter writer = new StreamWriter(_logFilePath, false))
            {
                writer.WriteLine($"=== Excel to CSV Converter Log ===");
                writer.WriteLine($"Started: {DateTime.Now}");
                writer.WriteLine("=====================================");
            }
        }

        public void Log(string message)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(_logFilePath, true))
                {
                    writer.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}");
                }
            }
            catch (Exception ex)
            {
                // Log internally if logging fails
                Console.WriteLine($"Error writing to log: {ex.Message}");
            }
        }
    }
}