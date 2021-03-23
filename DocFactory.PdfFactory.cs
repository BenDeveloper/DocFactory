using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using System.IO;

namespace DocFactory
{
    /// <summary>
    /// Partial class PDF document generation.
    /// Pdf can be get to File or Stream.
    /// </summary>
    public partial class DocFactory
    {
        /// <summary>
        /// Get document to Stream.
        /// </summary>
        /// <param name="landscapeMode"></param>
        /// <returns></returns>
        public Stream GetPDFToStream(bool landscapeMode = false)
        {
            Document doc = SetPdfDocument(landscapeMode);
            PdfDocumentRenderer pdfRenderer = GetPdfDocumentRenderer(doc);

            Stream stream = new MemoryStream();
            pdfRenderer.PdfDocument.Save(stream, false);

            return stream;
        }

        /// <summary>
        /// Save document to PDF.
        /// </summary>
        /// <param name="landscapeMode"></param>
        public void SavePDFToFile(bool landscapeMode = false)
        {
            Document doc = SetPdfDocument(landscapeMode);
            PdfDocumentRenderer pdfRenderer = GetPdfDocumentRenderer(doc);

            pdfRenderer.PdfDocument.Save(_fileInfo.FullName);
        }

        /// <summary>
        /// Get document renderer(Used for saving).
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private PdfDocumentRenderer GetPdfDocumentRenderer(Document document)
        {
            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(false)
            {
                Document = document
            };

            pdfRenderer.RenderDocument();

            return pdfRenderer;
        }

        /// <summary>
        /// PDF document generation.
        /// </summary>
        /// <param name="landscapeMode">Définir l'orientation en mode portrait.</param>
        /// <returns></returns>
        private Document SetPdfDocument(bool landscapeMode = false)
        {
            Document document = new Document();

            // Set landscape mode
            if (landscapeMode)
            {
                document.DefaultPageSetup.Orientation = Orientation.Landscape;
            }

            // Set Styles
            DefineStyles(document);
            document.AddSection();

            // Set title
            if (!string.IsNullOrEmpty(_title))
            {
                document.LastSection.AddParagraph(_title, "Heading1");
            }

            // Set Table
            Table table = new Table();
            table.Borders.Width = 0.75;

            // Set columns size and format
            int i = 0;
            foreach (ColumnDefinition c in _columns)
            {
                Column column = table.AddColumn(Unit.FromCentimeter(c.Size));
                column.Format.Alignment = ParagraphAlignment.Center;
            }

            // Set rows color
            Row row = table.AddRow();
            row.Shading.Color = Colors.LightSkyBlue;
            Cell cell;
            i = 0;

            // Fill column line, font size and bold
            foreach (ColumnDefinition c in _columns)
            {
                cell = row.Cells[i];
                Paragraph p = cell.AddParagraph(c.Name);
                p.Format.Font.Size = "10";
                p.Format.Font.Bold = true;
                i++;
            }

            // Set rows
            foreach (string[] r in _rows)
            {
                // Add row
                row = table.AddRow();

                // Fill row cells
                for (i = 0; i < _columns.Count; i++)
                {
                    // Fill cell and set font size
                    cell = row.Cells[i];
                    Paragraph p = cell.AddParagraph(r[i]);
                    p.Format.Font.Size = 8;
                }
            }

            table.SetEdge(0, 0, _columns.Count, _rows.Count + 1, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);

            document.LastSection.Add(table);

            return document;
        }

        /// <summary>
        /// Defines the PDF styles used in the document.
        /// </summary>
        public static void DefineStyles(Document document)
        {
            // Get the predefined style Normal.
            Style style = document.Styles["Normal"];

            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Verdana";

            // Heading1 to Heading9 are predefined styles with an outline level. An outline level
            // other than OutlineLevel.BodyText automatically creates the outline (or bookmarks) 
            // in PDF.
            style = document.Styles["Heading1"];
            style.Font.Size = 14;
            style.Font.Bold = true;
            style.Font.Color = Colors.Black;
            style.ParagraphFormat.PageBreakBefore = true;
            style.ParagraphFormat.SpaceAfter = 6;
        }
    }
}
