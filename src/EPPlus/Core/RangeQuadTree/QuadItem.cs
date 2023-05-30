using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace OfficeOpenXml.Core.RangeQuadTree
{
    internal enum IntersectType
    {
        OutSide = 0,
        Inside = 1,
        Partial = 2,
    }

    internal class QuadItem<T>
    {
        public QuadItem(QuadItem<T> parent, QuadRange dimension)
        {
            Parent = parent;
            Dimension = dimension;
        }
        public void Add(QuadRange range, T value)
        {
            if (!Dimension.IsMinimumSize)
            {
                if (Quads == null)
                {
                    Quads = new List<QuadItem<T>>();
                    AddQuads();
                }

                foreach (var q in Quads)
                {
                    var intersectType = q.Intersect(range);
                    if (intersectType == IntersectType.Inside)
                    {
                        q.Add(range, value);
                        return;
                    }
                    else if (intersectType == IntersectType.Partial)
                    {
                        break;
                    }
                }
            }
            Ranges.Add(new QuadRangeItem<T>(range, value));
        }

        public IntersectType Intersect(QuadRange range)
        {
            return Dimension.Intersect(range);
        }

        private void AddQuads()
        {
            var fr = Dimension.FromRow;
            var fc = Dimension.FromCol;
            var tr = Dimension.ToRow;
            var tc = Dimension.ToCol;
            var rows = tr - fr + 1;
            var cols = tc - fc + 1;
            if (rows / 2 > cols)
            {
                var qRows = rows / 4;
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc, fr + qRows - 1, tc)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr + qRows, fc, fr + qRows * 2 - 1, tc)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr + qRows * 2, fc, fr + qRows * 3 - 1, tc)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr + qRows * 3, fc, tr, tc)));
            }
            else if (cols / 2 > rows)
            {
                var qCols = cols / 4;
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc, tr, fc + qCols - 1)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc + qCols, tr, fc + qCols * 2 - 1)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc * qCols * 2, tr, fc + qCols * 3 - 1)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc * qCols * 3, tr, tc)));
            }
            else
            {
                var qRows = rows / 2;
                var qCols = cols / 2;
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc, fr + qRows - 1, fc + qCols - 1)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc + qCols, fr + qRows - 1, tc)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr + qRows, fc, tr, fc + qCols - 1)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr + qRows, fc + qCols, tr, tc)));
            }
        }

        public QuadRange Dimension { get; set; }
        public QuadItem<T> Parent { get; }
        /// <summary>
        /// Ranges intersecting with this quad.
        /// </summary>
        public List<QuadRangeItem<T>> Ranges { get; } = new List<QuadRangeItem<T>>();
        public List<QuadItem<T>> Quads { get; private set; }
        public override string ToString()
        {
            return Dimension.ToString();
        }

        internal void GetIntersectingRanges(QuadRange range, ref List<QuadRangeItem<T>> ranges)
        {
            if (Quads != null)
            {
                foreach (var q in Quads)
                {
                    if (q.Intersect(range) != IntersectType.OutSide)
                    {
                        q.GetIntersectingRanges(range, ref ranges);
                    }
                }
            }
            foreach(var r in Ranges)
            {
                if(r.Range.Intersect(range) != IntersectType.OutSide)
                {
                    ranges.Add(r);
                }
            }
        }
    }
}
