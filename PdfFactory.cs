using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using System.Collections.Generic;
using System.IO;

namespace ConsoleAppPDF
{
    /// <summary>
    /// Classe de gestion des PDF avec PDFSharp et MigraDoc.
    /// </summary>
    public class PdfFactory
    {
        /// <summary>
        /// Génère un PDF avec une table.
        /// </summary>
        /// <param name="columns">Liste des colonnes.</param>
        /// <param name="rows">Liste de tableaux de string des lignes.</param>
        /// <param name="title">(optionnel) Titre du tableau.</param>
        /// <returns></returns>
        public static Stream TableToPDF(List<ColumnDefinition> columns, List<string[]> rows, string title = null)
        {
            // Creation du document
            Document document = new Document();
            DefineStyles(document);
            document.AddSection();

            if (!string.IsNullOrEmpty(title))
            {
                document.LastSection.AddParagraph(title, "Heading1");
            }

            // Définit table
            Table table = new Table();
            table.Borders.Width = 0.75;

            // Définit les colonnes
            int i = 0;
            foreach (ColumnDefinition c in columns)
            {
                Column column = table.AddColumn(Unit.FromCentimeter(c.Size));
                column.Format.Alignment = ParagraphAlignment.Center;
            }

            // Définit la line des Noms de colonnes
            Row row = table.AddRow();
            row.Shading.Color = Colors.LightSkyBlue;
            Cell cell;
            i = 0;

            // Remplir cellules colonnes
            foreach (ColumnDefinition c in columns)
            {
                cell = row.Cells[i];
                Paragraph p = cell.AddParagraph(c.Name);
                p.Format.Font.Size = 10;
                p.Format.Font.Bold = true;
                i++;
            }

            // Créer les lignes
            foreach (string[] r in rows)
            {
                row = table.AddRow();
                // Remplir cellules colonnes
                for (i = 0; i < columns.Count; i++)
                {
                    cell = row.Cells[i];
                    Paragraph p = cell.AddParagraph(r[i]);
                    p.Format.Font.Size = 8;
                }
            }

            table.SetEdge(0, 0, columns.Count, rows.Count + 1, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);

            document.LastSection.Add(table);

            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(false)
            {
                Document = document
            };
            pdfRenderer.RenderDocument();

            Stream stream = new MemoryStream();
            pdfRenderer.PdfDocument.Save(stream, false);

            return stream;
        }

        /// <summary>
        /// Defines the styles used in the document.
        /// </summary>
        public static void DefineStyles(Document document)
        {
            // Get the predefined style Normal.
            MigraDoc.DocumentObjectModel.Style style = document.Styles["Normal"];
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

    /// <summary>
    /// Définit une colonne.
    /// </summary>
    public class ColumnDefinition
    {
        /// <summary>
        /// Column name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Column size in centimeters.
        /// </summary>
        public double Size { get; set; }
    }
}
