using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfSettings.PdfPageData
{
    internal class PdfExcelPageData
    {
        internal readonly int RowMaxCount;
        internal readonly int ColumnMaxCount;
        internal double contentHeight;
        internal double contentWidth;
        internal ExcelRangeBase PageRange;
        internal List<double[]> rowLineCoords = new List<double[]>();
        internal List<double[]> colLineCoords = new List<double[]>();

        public PdfExcelPageData(int rowCount, int columnCount)
        {
            this.RowMaxCount = rowCount;
            this.ColumnMaxCount = columnCount;
        }
    }
}
