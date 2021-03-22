using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using DocFactory;

namespace DocFactory
{
    /// <summary>
    /// This sample is the obligatory Hello World program.
    /// </summary>
    public class Program
    {
        static void Main(string[] args)
        {
            FileInfo file = new FileInfo("C:/temp/Test1.pdf");
            SetPDF(file.FullName);

            if (file.Exists)
            {
                Process.Start(file.FullName);
            }
        }

        private static void SetPDF(string fullName)
        {
            List<ColumnDefinition> c = new List<ColumnDefinition>
            {
                new ColumnDefinition() { Name = "Name", Size = 2.5 },
                new ColumnDefinition() { Name = "Number", Size = 2.5 },
                new ColumnDefinition() { Name = "Color", Size = 2.5 }
            };

            List<string[]> r = new List<string[]>();
            r.Add(new string[]
            {
                    "Tom", "1", "Blue" 
            });

            var f = new DocFactory(c, r, "Test Ben");
            Stream s = f.GetPDFStream();

            PdfSharp.Pdf.PdfDocument doc = new PdfSharp.Pdf.PdfDocument(s);
            doc.AddPage();

            doc.Save(s);
        }
    }
}