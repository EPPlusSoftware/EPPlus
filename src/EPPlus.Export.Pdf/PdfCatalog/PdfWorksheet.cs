using EPPlus.Export.Pdf.PdfLayout;
using OfficeOpenXml;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfWorksheet
    {
        public Dictionary<string, PdfCommentsAndNotes> CommentsAndNotesCollections = new Dictionary<string, PdfCommentsAndNotes>();
        public List<PdfRange>[] Ranges = null;
        public PdfHeaderFooterCollection HeaderFooters = null;

        public PdfRange CommentsAndNotes;

        //EPPlus references
        public ExcelWorksheet Worksheet { get; set; }
        public ExcelWorksheet CommentsAndNotesSheet { get; set; }
        public ExcelNamedStyleXml NormalStyle { get { return Worksheet.Workbook.Styles.GetNormalStyle(); } }

        //public PdfTextMap TextMap;
    }
}