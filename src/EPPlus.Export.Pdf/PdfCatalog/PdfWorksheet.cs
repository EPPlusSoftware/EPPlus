using EPPlus.Export.Pdf.PdfLayout;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfWorksheet
    {
        public Dictionary<string, PdfCommentsAndNotes> CommentsAndNotesCollections = null;
        public List<PdfRange>[] Ranges = null;
        public PdfHeaderFooterCollection HeaderFooters = null;
        public ExcelWorksheet Worksheet { get; set; }

        //public PdfTextMap TextMap;
    }
}
