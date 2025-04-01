using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Core.RangeQuadTree
{
    internal class QuadTreeAddress
    {
        internal static IList<QuadRangeItem<T>> ClearAddress<T>(QuadRangeItem<T> rangeItem, QuadRange clearedRange)
        {
            var r = rangeItem.Range; 
            var l=new List<QuadRangeItem<T>>();

            if(clearedRange.FromRow > r.FromRow && clearedRange.ToRow > r.FromRow)
            {
                l.Add(rangeItem.CloneWithNewAddress(r.FromRow, r.FromCol, clearedRange.FromRow-1, r.ToRow));
            }

            if(clearedRange.FromCol > r.FromCol && clearedRange.ToCol > r.FromCol)
            {
                l.Add(rangeItem.CloneWithNewAddress(Math.Max(r.FromRow, clearedRange.FromRow), r.FromCol, r.ToRow, clearedRange.FromCol-1));
            }

            if(clearedRange.ToCol < r.ToCol)
            {
                l.Add(rangeItem.CloneWithNewAddress(Math.Max(r.FromRow, clearedRange.FromRow), clearedRange.ToCol+1, r.ToRow, r.ToCol));
            }

            if(clearedRange.ToRow < r.ToRow)
            {
                l.Add(rangeItem.CloneWithNewAddress(clearedRange.ToRow+1, Math.Max(r.FromCol, clearedRange.FromCol), r.ToRow, Math.Min(r.ToCol, clearedRange.ToCol)));
            }
            //if (clearedRange.FromRow >= r.FromRow)
            //{
            //    if (clearedRange.FromCol > r.FromCol)
            //    {
            //        l.Add(rangeItem.CloneWithNewAddress(r.FromRow, r.FromCol, r.ToRow, clearedRange.FromCol-1));
            //    }
            //    if (clearedRange.ToCol < r.ToCol)
            //    {
            //        l.Add(rangeItem.CloneWithNewAddress(r.FromRow, clearedRange.FromCol+1, r.ToRow, r.ToCol));
            //    }
            //}
            //else
            //{
            //    int fromRow;
            //    if (clearedRange.FromCol <= r.FromCol && clearedRange.ToCol < r.ToCol)
            //    {
            //        l.Add(rangeItem.CloneWithNewAddress(r.FromRow, r.FromCol, clearedRange.FromRow - 1, clearedRange.ToCol));
            //        fromRow = clearedRange.FromRow - 1;
            //    }
            //    else
            //    {
            //        fromRow=r.FromRow;
            //    }
            //    if (clearedRange.FromCol > r.FromCol)
            //    {
            //        l.Add(rangeItem.CloneWithNewAddress(fromRow, r.FromCol, r.ToRow, clearedRange.FromCol - 1));
            //    }
            //    if (clearedRange.ToCol < r.ToCol)
            //    {
            //        l.Add(rangeItem.CloneWithNewAddress(r.FromRow, clearedRange.ToCol + 1, r.ToRow, r.ToCol));
            //    }
            //}

            //if(clearedRange.ToRow < r.ToRow)
            //{
            //    l.Add(rangeItem.CloneWithNewAddress(clearedRange.ToRow + 1, Math.Max(clearedRange.FromCol, r.FromCol), r.ToRow, Math.Min(clearedRange.ToCol, r.ToCol)));
            //}
            return l;
        }
    }
}