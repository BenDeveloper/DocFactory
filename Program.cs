using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using ConsoleAppPDF;

namespace HelloWorld
{
    /// <summary>
    /// This sample is the obligatory Hello World program.
    /// </summary>
    public class Program
    {
        static void Main(string[] args)
        {
            FileInfo file = new FileInfo("C:/temp/Test1.pdf");
            CreatePDFwithMigraDoc(file.FullName);

            if (file.Exists)
            {
                Process.Start(file.FullName);
            }
        }

        private static void CreatePDFwithMigraDoc(string fullName)
        {
            // Columns list
            List<ColumnDefinition> c = new List<ColumnDefinition>();
            c.Add(new ColumnDefinition() { Name = "Name", Size = 2.5 });
            c.Add(new ColumnDefinition() { Name = "number", Size = 2.5 });
            c.Add(new ColumnDefinition() { Name = "Color", Size = 2.5 });

            // Row list
            List<string[]> r = new List<string[]>();
            r.Add(new string[]
            {
                    "Tom", "1", "Blue" 
            });

            // Create table
            Stream s = PdfFactory.TableToPDF(c, r, "Test Ben");

            PdfSharp.Pdf.PdfDocument doc = new PdfSharp.Pdf.PdfDocument(s);
            doc.AddPage();

            doc.Save(fullName);
        }
    }
}