using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.RangeQuadTree
{
    internal class QuadTree<T>
    {
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

        internal List<QuadRangeItem<T>> GetIntersectingRanges(QuadRange range)
        {
            var ranges = new List<QuadRangeItem<T>>();
            Root.GetIntersectingRanges(range, ref ranges);
            return ranges;
        }
    }
}
