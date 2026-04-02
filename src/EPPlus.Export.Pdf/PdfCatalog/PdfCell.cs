using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Fonts.OpenType.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal struct PdfCell
    {
        public PdfCellStyle CellStyle;
        public PdfCellAlignmentData ContentAligmnet;
        public List<PdfTextFormat> TextFormats;
        public TextLayoutEngine TextLayoutEngine { get; set; }




        //double x, y, width, height;
        //bool isMerged;
        //string address;
        //int row, col;
        //Fill
        //Text
        //Picute
        //Border


    }
}
