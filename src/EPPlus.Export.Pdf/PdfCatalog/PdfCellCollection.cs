using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfCellCollection
    {
        private PdfCell[,] Cells;

        public readonly int FromRow;
        public readonly int FromColumn;
        public readonly int ToRow;
        public readonly int ToColumn;

        public PdfCellCollection(int fromRow, int toRow, int fromColumn, int toColumn)
        {
            if(fromRow > toRow) throw new ArgumentOutOfRangeException("Invalid row range. toRow must be equal or greater than fromRow");
            if(fromColumn > toColumn) throw new ArgumentOutOfRangeException("Invalid column range. toColumn must be equal or greater than fromColumn");
            FromRow = fromRow;
            FromColumn = fromColumn;
            ToRow = toRow;
            ToColumn = toColumn;
            int x = toRow - fromRow + 1;
            int y = toColumn - fromColumn + 1;

            Cells = new PdfCell[x, y];
        }

        public PdfCell this[int row, int column]
        {
            get
            {
                var r = row - FromRow;
                var c = column - FromColumn;
                if (r < 0 || c < 0 || r >= Cells.GetLength(0) || c >= Cells.GetLength(1))
                {
                    return null;
                }
                return Cells[r, c];
            }
            set
            {
                Cells[row - FromRow, column - FromColumn] = value;
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
