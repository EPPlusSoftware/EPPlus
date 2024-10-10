using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.RangeQuadTree
{
    internal class QuadTree<T>
    {

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
            Root.Add(quadRange, item);
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
            Root.InsertRow(fromRow, rows);
        }
    }
}
