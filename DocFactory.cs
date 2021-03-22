using MigraDoc.DocumentObjectModel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DocFactory
{
    /// <summary>
    /// Classe de gestion des PDF avec PDFSharp et MigraDoc.
    /// </summary>
    public partial class DocFactory
    {
        private string _title;
        private readonly List<ColumnDefinition> _columns;
        private readonly List<string[]> _rows;

        /// <summary>
        /// Constructeur.
        /// </summary>
        /// <param name="columns"></param>
        /// <param name="rows"></param>
        /// <param name="title"></param>
        public DocFactory(List<ColumnDefinition> columns, List<string[]> rows, string title = null)
        {
            _columns = columns ?? throw new ArgumentNullException(nameof(columns));
            _rows = rows ?? throw new ArgumentNullException(nameof(rows));

            if (!_rows.Any())
            {
                throw new ArgumentOutOfRangeException(nameof(rows));
            }

            _title = title;
        }

        /// <summary>
        /// Defines the styles used in the document.
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
