using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.PdfExport
{
    public class PdfExportSettingsOverrides
    {
        /// <summary>
        /// Prefix that marks a defined name as a PDF bookmark. The prefix is stripped from the
        /// bookmark title. The bookmark text is taken from the target cell's content, or the defined
        /// name (minus the prefix) when the cell is empty. Default is <c>"bm_"</c>.
        /// This is the document-wide default; it can be overridden per worksheet.
        /// </summary>
        public string BookmarkNamePrefix { get; set; }
    }
}
