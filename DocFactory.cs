using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocFactory
{
    /// <summary>
    /// Partial class for commun constructor and methods for document generation.
    /// Current document generated : PDF(Stream, File).
    /// </summary>
    public partial class DocFactory
    {
        private string _title;
        private readonly List<ColumnDefinition> _columns;
        private readonly List<string[]> _rows;
        private readonly FileInfo _fileInfo;

        /// <summary>
        /// Constructeur.
        /// </summary>
        /// <param name="columns"></param>
        /// <param name="rows"></param>
        /// <param name="title"></param>
        /// <param name="fileInfo"></param>
        public DocFactory(List<ColumnDefinition> columns, List<string[]> rows, string title = null, FileInfo fileInfo = null)
        {
            _columns = columns ?? throw new ArgumentNullException(nameof(columns));
            _rows = rows ?? throw new ArgumentNullException(nameof(rows));

            if (!_rows.Any())
            {
                throw new ArgumentOutOfRangeException(nameof(rows));
            }

            _title = title;
            _fileInfo = fileInfo;
        }
    }
}
