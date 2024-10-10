using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;

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

        internal void InsertRow(int fromRow, int rows)
        {
            if (ToRow < fromRow) return;
            if (FromRow >= fromRow)
            {
                FromRow += rows;
            }
            ToRow += rows;
        }
        internal void InsertColumn(int fromCol, int cols)
        {
            if (ToCol < fromCol) return;
            if (FromCol >= fromCol)
            {
                FromCol += cols;
            }
            ToCol += cols;
        }
        internal void DeleteRow(int fromRow, int rows)
        {
            if (ToRow < fromRow) return;
            if (FromRow >= fromRow)
            {
                FromRow -= rows;
            }
            ToRow -= rows;
        }
        internal void DeleteColumn(int fromCol, int cols)
        {
            if (ToCol < fromCol) return;
            if (FromCol >= fromCol)
            {
                FromCol -= cols;
            }
            ToCol -= cols;
        }

    }
}