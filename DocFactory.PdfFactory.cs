using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocFactory
{
    public partial class DocFactory
    {
        /// <summary>
        /// Génère un Stream de document PDF.
        /// </summary>
        /// <param name="landscapeMode">Définir l'orientation en mode portrait.</param>
        /// <returns></returns>
        public Stream GetPDFStream(bool landscapeMode = false)
        {
            // Creation du document
            Document document = new Document();

            if (landscapeMode)
            {
                document.DefaultPageSetup.Orientation = Orientation.Landscape;
            }

            DefineStyles(document);
            document.AddSection();

            if (!string.IsNullOrEmpty(_title))
            {
                document.LastSection.AddParagraph(_title, "Heading1");
            }

            // Définit table
            Table table = new Table();
            table.Borders.Width = 0.75;

            // Définit les colonnes
            int i = 0;
            foreach (ColumnDefinition c in _columns)
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
            foreach (ColumnDefinition c in _columns)
            {
                cell = row.Cells[i];
                Paragraph p = cell.AddParagraph(c.Name);
                p.Format.Font.Size = "10";
                p.Format.Font.Bold = true;
                i++;
            }

            // Créer les lignes
            foreach (string[] r in _rows)
            {
                row = table.AddRow();
                // Remplir cellules colonnes
                for (i = 0; i < _columns.Count; i++)
                {
                    // Ajoute Paragraphe et définit la taille
                    cell = row.Cells[i];
                    Paragraph p = cell.AddParagraph(r[i]);
                    p.Format.Font.Size = 8;
                }
            }

            table.SetEdge(0, 0, _columns.Count, _rows.Count + 1, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);

            document.LastSection.Add(table);

            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(false);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();

            Stream stream = new MemoryStream();
            pdfRenderer.PdfDocument.Save(stream);

            return stream;
        }
    }
}
