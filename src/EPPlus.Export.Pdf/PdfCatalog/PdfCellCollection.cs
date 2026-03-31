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

        private readonly int fromRow;
        private readonly int fromColumn;

        public PdfCellCollection(int fromRow, int toRow, int fromColumn, int toColumn)
        {
            if(fromRow > toRow) throw new ArgumentOutOfRangeException("Invalid row range. toRow must be equal or greater than fromRow");
            if(fromColumn > toColumn) throw new ArgumentOutOfRangeException("Invalid column range. toColumn must be equal or greater than fromColumn");
            this.fromRow = fromRow;
            this.fromColumn = fromColumn;
            int x = toRow - fromRow + 1;
            int y = toColumn - fromColumn + 1;

            Cells = new PdfCell[x, y];
        }

        public PdfCell this[int row, int column]
        {
            get
            {
                return Cells[row - fromRow, column - fromColumn];
            }
            set
            {
                Cells[row - fromRow, column - fromColumn] = value;
            }
        }

        public PdfCell GetInternal(int row, int column)
        {
            return Cells[row, column];
        }

        public void SetInternal(int row, int column, PdfCell value)
        {
            Cells[row, column] = value;
        }
    }
}
