using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
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

        public void Add(QuadRange quadRange, T item)
        {
            if (Root.Dimension.IsOutsideOfBounds(quadRange))
            {
                ExpandTree(quadRange);
            }

            Root.Add(quadRange, item);
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

                if (d.ToCol == ExcelPackage.MaxColumns)
                {
                    quadItem = new QuadItem<T>(null, new QuadRange(d.FromRow, d.FromCol, d.ToRow * 4, d.ToCol));
                }
                else if (d.ToRow == ExcelPackage.MaxRows)
                {
                    quadItem = new QuadItem<T>(null, new QuadRange(d.FromRow, d.FromCol, d.ToRow, d.ToCol * 4));
                }
                else
                {
                    quadItem = new QuadItem<T>(null, new QuadRange(d.FromRow, d.FromCol, d.ToRow * 2, d.ToCol * 2));
                }

                if(quadItem.Dimension.ToRow>ExcelPackage.MaxRows)
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
        public void InsertRow(int fromRow, int rows)
        {
            if (fromRow <= Root.Dimension.ToRow)
            {
                var cr = new QuadRange(1, 1, Root.Dimension.ToRow + rows, Root.Dimension.ToCol);
                if (Root.Dimension.IsOutsideOfBounds(cr))
                {
                    ExpandTree(cr);
                }
            }

            Root.InsertRow(fromRow, rows);
        }
        public void InsertColumn(int fromCol, int cols)
        {
            if (fromCol <= Root.Dimension.ToCol)
            {
                var cr = new QuadRange(1, 1, Root.Dimension.ToRow, Root.Dimension.ToCol + cols);
                if (Root.Dimension.IsOutsideOfBounds(cr))
                {
                    ExpandTree(cr);
                }
            }

            Root.InsertCol(fromCol, cols);
        }
        public void DeleteRow(int fromRow, int rows)
        {
            Root.DeleteRow(fromRow, rows);
        }
        public void DeleteCol(int fromCol, int cols)
        {
            Root.DeleteCol(fromCol, cols);
        }
        public void Clear(int fromRow, int fromCol, int toRow, int toCol)
        {
            var range=new QuadRange(fromRow, fromCol, toRow, toCol);
            Root.Clear(range);
        }
    }
}
