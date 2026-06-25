/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using System;

namespace OfficeOpenXml.Export.PdfExport.Data
{
    internal class PdfCellCollection
    {
        public readonly int FromRow;
        public readonly int FromColumn;
        public readonly int ToRow;
        public readonly int ToColumn;

        private PdfCell[,] Cells;

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
