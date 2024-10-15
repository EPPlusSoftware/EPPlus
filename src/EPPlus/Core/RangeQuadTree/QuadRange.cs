using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq.Expressions;

namespace OfficeOpenXml.Core.RangeQuadTree
{
    internal class QuadRange
    {
        public const int MinSize = 30;
        public int FromRow { get; set; }
        public int FromCol { get; set; }
        public int ToRow { get; set; }
        public int ToCol { get; set; }
        public bool IsMinimumSize
        {
            get
            {
                return ToRow - FromRow < MinSize &&
                    ToCol - FromCol < MinSize;
            }
        }
        public QuadRange(FormulaRangeAddress range) : this(range.FromRow, range.FromCol, range.ToRow, range.ToCol)
        {

        }
        public QuadRange(ExcelAddressBase address) : this(address._fromRow, address._fromCol, address._toRow, address._toCol)
        {

        }
        public QuadRange(int fromRow, int fromCol, int toRow, int toCol)
        {
            FromRow = fromRow;
            FromCol = fromCol;
            ToRow = toRow;
            ToCol = toCol;
        }

        public override string ToString()
        {
            return ExcelCellBase.GetAddress(FromRow, FromCol, ToRow, ToCol);
        }

        internal IntersectType Intersect(QuadRange range)
        {
            if (range.FromRow >= FromRow && range.ToRow <= ToRow &&
               range.FromCol >= FromCol && range.ToCol <= ToCol)
            {
                return IntersectType.Inside;
            }

            if (range.FromRow <= ToRow && range.FromCol <= ToCol
                   &&
                   FromRow <= range.ToRow && FromCol <= range.ToCol)
            {
                return IntersectType.Partial;
            }
            return IntersectType.OutSide;
        }

        /// <summary>
        /// Inserts row(s) into the quad tree. Intersecting ranges will be expanded. Ranges after the row will be shifted.
        /// </summary>
        /// <param name="fromRow">Insert from this row.</param>
        /// <param name="rows">Insert this number of cols.</param>
        /// <returns>True if affected</returns>
        internal bool InsertRow(int fromRow, int rows)
        {
            if (ToRow < fromRow) return false;
            if (FromRow >= fromRow)
            {
                FromRow += rows;
            }
            ToRow += rows;
            return true;
        }
        /// <summary>
        /// Inserts columns(s) into the quad tree. Intersecting ranges will be expanded. Ranges after the column will be shifted.
        /// </summary>
        /// <param name="fromCol">Insert from this column.</param>
        /// <param name="cols">Insert this number of column.</param>
        /// <returns>True if affected</returns>        
        internal bool InsertColumn(int fromCol, int cols)
        {
            if (ToCol < fromCol) return false;
            if (FromCol >= fromCol)
            {
                FromCol += cols;
            }
            ToCol += cols;
            return true;
        }
        internal bool DeleteRow(int fromRow, int rows)
        {
            if (ToRow < fromRow) return false;

            var toRow = fromRow + rows-1;
            if (fromRow <= FromRow && toRow >= ToRow) //Within, delete
            {
                FromRow = -1;
                ToRow =-1;
                return true;
            }
            
            if(FromRow >= fromRow+rows) //Adjust FromRow/ToRow
            {
                FromRow -= rows;
            }
            else if(FromRow > fromRow)
            {
                FromRow=fromRow;
            }

            if (ToRow >= fromRow + rows) //Adjust FromRow/ToRow
            {
                ToRow -= rows;
            }
            else if (ToRow > fromRow)
            {
                ToRow = fromRow;
            }

            return true;
            
        }
        internal bool DeleteColumn(int fromCol, int cols)
        {
            if (ToCol < fromCol) return false;

            var toCol = fromCol + cols - 1;
            if (fromCol <= FromCol && toCol >= ToCol) //Within, delete
            {
                FromCol = -1;
                ToCol = -1;
                return true;
            }

            if (FromCol >= fromCol + cols) //Adjust FromRow/ToRow
            {
                FromCol -= cols;
            }
            else if (FromCol > fromCol)
            {
                FromCol = fromCol;
            }

            if (ToCol >= fromCol + cols) //Adjust FromRow/ToRow
            {
                ToCol -= cols;
            }
            else if (ToCol > fromCol)
            {
                ToCol = fromCol;
            }

            return true;
        }

        internal bool IsOutsideOfBounds(QuadRange range)
        {
            return ToRow < range.ToRow || ToCol < range.ToCol;
        }

        internal bool IsInside(QuadRange r)
        {
            return 
               FromRow <= r.FromRow &&
               FromCol <= r.FromCol &&
               ToRow >= r.ToRow &&
               ToCol >= r.ToCol;
        }
        
    }
}