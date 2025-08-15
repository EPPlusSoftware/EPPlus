using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Diagnostics;

namespace OfficeOpenXml.Core.RangeQuadTree
{
    [DebuggerDisplay("{Range},{Value}")]
    internal struct QuadRangeItem<T>
    {
        public QuadRangeItem(QuadRange range, T value)
        {
            Range=range;
            Value=value;
        }
        public QuadRange Range{ get; }
        public T Value { get; }
        internal QuadRangeItem<T> CloneWithNewAddress(int fromRow, int fromCol, int toRow, int toCol)
        {
            var r = new QuadRange(fromRow, fromCol, toRow, toCol);
            return new QuadRangeItem<T>(r, Value);
        }
    }
}