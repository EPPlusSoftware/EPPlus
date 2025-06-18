using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfSettings.FontData
{
    internal class PdfExcelFontData
    {
        internal readonly int RowCount;
        internal readonly int ColumnCount;
        internal readonly double FontSize;
        internal readonly double paddingAdd;

        public PdfExcelFontData(int rowCount, int columnCount, double fontSize)
        {
            this.RowCount = rowCount;
            this.ColumnCount = columnCount;
            this.FontSize = fontSize;
        }
    }
}
