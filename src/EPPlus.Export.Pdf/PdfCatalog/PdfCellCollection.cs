using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfCellCollection
    {
        internal PdfCell[,] Cells;

        public PdfCellCollection(int rows, int column)
        {
            Cells = new PdfCell[rows, column];
        }

        public void AddCell(int row, int column)
        {
            Cells[row - 1, column - 1] = new PdfCell();
        }

        public PdfCell GetCell(int row, int column)
        {
            return Cells[row - 1, column - 1];
        }
    }
}
