using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Graphics;
using OfficeOpenXml;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfCell
    {
        public bool Hidden;
        public PdfCellStyle CellStyle;
        public PdfCellAlignmentData ContentAligmnet;
        public List<PdfTextFormat> TextFormats { get; set; }
        public double ColumnWidth { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }

        public TextLayoutEngine TextLayoutEngine { get; set; }

        public bool Merged;
        public PdfCell Main;
        public ExcelAddressBase MergedAddress;



        //public int FromCol;
        //public int ToCol;
        //public int FromRow;
        //public int ToRow;
        //double x, y, width, height;
        //string address;
        //int row, col;
        //Picute
        //kolla på andra map och gör denna lik så att vi kan avnänada samma kod till gridlines 

    }
}
