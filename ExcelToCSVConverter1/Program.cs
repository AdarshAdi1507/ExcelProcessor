using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Data;
using System.Diagnostics;
using System.Collections.Generic;
using ExcelToCSVConverter1.Forms;

namespace ExcelToCSVConverter1
{
    // Main application entry point
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Console.WriteLine("Hello world");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
