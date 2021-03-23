using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace DocFactory
{
    /// <summary>
    /// Console test program to test PDF generation.
    /// </summary>
    public class Program
    {
        static void Main(string[] args)
        {
            FileInfo file = new FileInfo("C:/temp/TestPDF.pdf");

            SetPDF(file);

            if (file.Exists)
            {
                Process.Start(file.FullName);
            }
        }

        /// <summary>
        /// Set columns, rows, title(facultative), and file info(facultative).
        /// </summary>
        /// <param name="fileInfo"></param>
        private static void SetPDF(FileInfo fileInfo)
        {
            // Set Columns
            List<ColumnDefinition> c = new List<ColumnDefinition>
            {
                new ColumnDefinition() { Name = "Name", Size = 2.5 },
                new ColumnDefinition() { Name = "Number", Size = 2.5 },
                new ColumnDefinition() { Name = "Color", Size = 2.5 }
            };

            // Set rows
            List<string[]> r = new List<string[]>()
            {
                new string[] { "Tom", "1", "Blue" },
                new string[] { "Ben", "2", "Green" },
                new string[] { "Moon", "3", "Black" },
            };

            var f = new DocFactory(c, r, "PDF Table Title", fileInfo);

            f.SavePDFToFile();
        }
    }
}