using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.RangeQuadTree
{
    internal class QuadTree<T>
    {
        public QuadTree() : this(1, 1, QuadRange.MinSize, QuadRange.MinSize)
        {

        }
        public QuadTree(FormulaRangeAddress range) : this(range.FromRow, range.FromCol, range.ToRow, range.ToCol)
        {

        }
        public QuadTree(ExcelAddressBase address) : this(address._fromRow, address._fromCol, address._toRow, address._toCol)
        {
            
        }

        public QuadTree(int fromRow, int fromCol, int toRow, int toCol)
        {
            var dimension = new QuadRange(fromRow, fromCol, toRow, toCol);
            Root = new QuadItem<T>(null, dimension);            
        }

        public QuadItem<T> Root { get; private set; }
        /// <summary>
        /// The maximum column of all ranges in the tree. -1 if not ranges exist.
        /// </summary>
        public int MaxCol { get; private set; }
        /// <summary>
        /// The maximum row of all ranges in the tree. -1 if not ranges exist.
        /// </summary>
        public int MaxRow { get; private set; }

        public void Add(QuadRange quadRange, T item)
        {
            if (Root.Dimension.IsOutsideOfBounds(quadRange))
            {
                ExpandTree(quadRange);
            }

            Root.Add(quadRange, item);

            if (quadRange.ToRow > MaxRow)
            { 
                MaxRow = quadRange.ToRow; 
            }

            if(quadRange.ToCol > MaxCol)
            {
                MaxCol = quadRange.ToCol;
            }
        }
        private void ExpandTree(QuadRange range)
        {
            var d = Root.Dimension;
            do
            {
                QuadItem<T> quadItem;

                if (d.ToRow > ExcelPackage.MaxRows || d.ToCol > ExcelPackage.MaxColumns)
                {
                    throw new InvalidOperationException($"Quad tree can not expand over worksheet max limits (rows:{ExcelPackage.MaxRows}, columns {ExcelPackage.MaxColumns})");
                }

                if (d.ToCol == ExcelPackage.MaxColumns || (d.ToCol * 2) < MaxCol)
                {
                    quadItem = new QuadItem<T>(null, new QuadRange(d.FromRow, d.FromCol, d.ToRow * 4, d.ToCol));
                }
                else if (d.ToRow == ExcelPackage.MaxRows || (d.ToRow * 2) < MaxRow)
                {
                    quadItem = new QuadItem<T>(null, new QuadRange(d.FromRow, d.FromCol, d.ToRow, d.ToCol * 4));
                }
                else
                {
                    quadItem = new QuadItem<T>(null, new QuadRange(d.FromRow, d.FromCol, d.ToRow * 2, d.ToCol * 2));
                }

                if(quadItem.Dimension.ToRow > ExcelPackage.MaxRows)
                {
                   quadItem.Dimension.ToRow = ExcelPackage.MaxRows;
                }
                if(quadItem.Dimension.ToCol>ExcelPackage.MaxColumns)
                {
                   quadItem.Dimension.ToCol = ExcelPackage.MaxColumns;
                }

                quadItem.AddQuads(Root);
                Root.Parent = quadItem;
                Root = quadItem;
                d = quadItem.Dimension;
            }
            while (d.IsOutsideOfBounds(range));
        }

        internal List<QuadRangeItem<T>> GetIntersectingRangeItems(QuadRange range)
        {
            var ranges = new List<QuadRangeItem<T>>();
            Root.GetIntersectingRangeItems(range, ref ranges);
            return ranges;
        }
        internal List<QuadRange> GetIntersectingRanges(QuadRange range)
        {
            var ranges = new List<QuadRange>();
            Root.GetIntersectingRanges(range, ref ranges);
            return ranges;
        }
        public void InsertRow(int fromRow, int rows, int fromCol=1, int toCol = ExcelPackage.MaxColumns)
        {
            Root.InsertRow(fromRow, rows, fromCol, toCol, out bool overflow);
            if(overflow)
            {
                var cr = new QuadRange(1, 1, Root.Dimension.ToRow + rows, Root.Dimension.ToCol);
                ExpandTree(cr);
            }
        }
        public void InsertColumn(int fromCol, int cols, int fromRow=1, int toRow = ExcelPackage.MaxRows)
        {
            Root.InsertCol(fromCol, cols, fromRow, toRow, out bool overflow);
            if(overflow)
            {
                var cr = new QuadRange(1, 1, Root.Dimension.ToRow, Root.Dimension.ToCol + cols);
                ExpandTree(cr);
            }
        }
        public void DeleteRow(int fromRow, int rows, int fromCol = 1, int toCol = ExcelPackage.MaxRows)
        {
            Root.DeleteRow(fromRow, rows, fromCol, toCol);
        }
        public void DeleteCol(int fromCol, int cols, int fromRow = 1, int toRow = ExcelPackage.MaxRows)
        {
            Root.DeleteCol(fromCol, cols, fromRow, toRow);
        }
        internal void Clear(ExcelAddressBase a, T item)
        {
            Clear(a._fromRow, a._fromCol, a._toRow, a._toCol, item);
        }
        public void Clear(int fromRow, int fromCol, int toRow, int toCol)
        {
            var range = new QuadRange(fromRow, fromCol, toRow, toCol);
            Root.Clear(range, null);
        }
        public void Clear(int fromRow, int fromCol, int toRow, int toCol, T item)
        {
            var range=new QuadRange(fromRow, fromCol, toRow, toCol);
            Root.Clear(range, item);
        }

        internal void UpdateAddress(ExcelAddress addressToClear, ExcelAddress addressToAdd, T item)
        {
            if (addressToClear != null)
            {
                if (addressToClear.Addresses == null)
                {
                    Clear(addressToClear._fromRow, addressToClear._fromCol, addressToClear._toRow, addressToClear._toCol, item);
                }
                else
                {
                    foreach (var a in addressToClear.Addresses)
                    {
                        Clear(a._fromRow, a._fromCol, a._toRow, a._toCol, item);
                    }
                }
            }

            if(addressToAdd != null)
            {
                if (addressToAdd.Addresses == null)
                {
                    Add(new QuadRange(addressToAdd), item);
                }
                else
                {
                    foreach (var a in addressToAdd.Addresses)
                    {
                        Add(new QuadRange(a), item);
                    }
                }
            }           
        }
    }
}
