using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace OfficeOpenXml.Core.RangeQuadTree
{
    internal class QuadItem<T>
    {
        public QuadItem(QuadItem<T> parent, QuadRange dimension)
        {
            Parent = parent;
            Dimension = dimension;
        }
        public void Add(QuadRange range, T value)
        {
            Add(new QuadRangeItem<T>(range, value));
        }
        public void Add(QuadRangeItem<T> item)
        {
            if (!Dimension.IsMinimumSize)
            {
                if (Quads == null)
                {
                    AddQuads();
                }

                foreach (var q in Quads)
                {
                    var intersectType = q.Intersect(item.Range);
                    if (intersectType == IntersectType.Inside)
                    {
                        q.Add(item);
                        return;
                    }
                    else if (intersectType == IntersectType.Partial)
                    {
                        break;
                    }
                }
            }
            Ranges.Add(item);
        }

        public IntersectType Intersect(QuadRange range)
        {
            return Dimension.Intersect(range);
        }

        internal void AddQuads(QuadItem<T> firstItem=null)
        {
            var fr = Dimension.FromRow;
            var fc = Dimension.FromCol;
            var tr = Dimension.ToRow;
            var tc = Dimension.ToCol;
            var rows = tr - fr + 1;
            var cols = tc - fc + 1;
            Quads = new List<QuadItem<T>>();
            if (firstItem!=null)
            {
                Quads.Add(firstItem);
            }
            if (rows / 2 > cols)
            {
                var qRows = rows / 4;
                if(firstItem == null) Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc, fr + qRows - 1, tc)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr + qRows, fc, fr + qRows * 2 - 1, tc)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr + qRows * 2, fc, fr + qRows * 3 - 1, tc)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr + qRows * 3, fc, tr, tc)));
            }
            else if (cols / 2 > rows)
            {
                var qCols = cols / 4;
                if (firstItem == null) Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc, tr, fc + qCols - 1)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc + qCols, tr, fc + qCols * 2 - 1)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc * qCols * 2, tr, fc + qCols * 3 - 1)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc * qCols * 3, tr, tc)));
            }
            else
            {
                var qRows = rows / 2;
                var qCols = cols / 2;
                if (firstItem == null) Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc, fr + qRows - 1, fc + qCols - 1)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr, fc + qCols, fr + qRows - 1, tc)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr + qRows, fc, tr, fc + qCols - 1)));
                Quads.Add(new QuadItem<T>(this, new QuadRange(fr + qRows, fc + qCols, tr, tc)));
            }
        }

        public QuadRange Dimension { get; set; }
        public QuadItem<T> Parent { get; internal set; }
        /// <summary>
        /// Ranges intersecting with this quad.
        /// </summary>
        public List<QuadRangeItem<T>> Ranges { get; } = new List<QuadRangeItem<T>>();
        public List<QuadItem<T>> Quads { get; private set; }
        public override string ToString()
        {
            return Dimension.ToString();
        }

        internal void GetIntersectingRangeItems(QuadRange range, ref List<QuadRangeItem<T>> ranges)
        {
            if (Quads != null)
            {
                foreach (var q in Quads)
                {
                    if (q.Intersect(range) != IntersectType.OutSide)
                    {
                        q.GetIntersectingRangeItems(range, ref ranges);
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
        internal void GetIntersectingRanges(QuadRange range, ref List<QuadRange> ranges)
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
            foreach (var r in Ranges)
            {
                if (r.Range.Intersect(range) != IntersectType.OutSide)
                {
                    ranges.Add(r.Range);
                }
            }
        }

        internal bool DeleteRow(int fromRow, int rows, int fromCol, int toCol)
        {
            if(fromRow > Dimension.ToRow)
            {
                return false;
            }

            var qadd = new Dictionary<int, QuadRangeItem<T>>();
            var l = new List<QuadRangeItem<T>>();
            var ret = false;
            for (int i = 0; i < Ranges.Count; i++)
            {
                var r = Ranges[i];
                if (r.Range.FromCol < fromCol)
                {
                    l.Add(
                        r.CloneWithNewAddress(r.Range.FromRow, r.Range.FromCol, r.Range.ToRow, fromCol - 1));
                    r.Range.FromCol = fromCol;
                }
                if (r.Range.ToCol > toCol)
                {
                    l.Add(
                        r.CloneWithNewAddress(r.Range.FromRow, toCol + 1, r.Range.ToRow, r.Range.ToCol));
                    r.Range.ToCol = toCol;
                }

                if (r.Range.DeleteRow(fromRow, rows))
                {
                    ret = true;
                    if (r.Range.FromRow<1)
                    {
                        Ranges.RemoveAt(i);
                        i--;
                        continue;
                    }
                    if (r.Range.FromRow < Dimension.FromRow)
                    {
                        AddParentDeleteRow(i, r);
                        i--;
                    }
                    else if(Quads!=null && FitsInQuad(r, out int quadIx))
                    {
                        qadd.Add(quadIx, r);
                        Ranges.RemoveAt(i--);
                    }
                }
            }

            if (Quads != null)
            {
                foreach (var q in Quads)
                {
                    if (q.DeleteRow(fromRow, rows, fromCol, toCol))
                    {
                        ret = true;
                    }
                }
            }

            Ranges.AddRange(l);
            foreach (var q in qadd)
            {
                Quads[q.Key].Add(q.Value);
            }
            return ret;
        }
        internal bool DeleteCol(int fromCol, int cols, int fromRow = 1, int toRow = ExcelPackage.MaxRows)
        {
            if (fromCol > Dimension.ToCol)
            {
                return false;
            }

            var qadd = new Dictionary<int, QuadRangeItem<T>>();
            var l = new List<QuadRangeItem<T>>();
            var ret = false;
            for (int i = 0; i < Ranges.Count; i++)
            {
                var r = Ranges[i];
                if (r.Range.FromRow < fromRow)
                {
                    l.Add(
                        r.CloneWithNewAddress(r.Range.FromRow, r.Range.FromCol, fromRow - 1, r.Range.ToCol));
                    r.Range.FromRow = fromRow;
                }

                if (r.Range.ToRow > toRow)
                {
                    l.Add(
                        r.CloneWithNewAddress(toRow + 1, r.Range.FromCol, r.Range.ToRow, r.Range.ToCol));
                    r.Range.ToRow = toRow;
                }

                if (r.Range.DeleteColumn(fromCol, cols))
                {
                    ret = true;
                    if (r.Range.FromCol < 1)
                    {
                        Ranges.RemoveAt(i);
                        i--;
                        continue;
                    }
                    if (r.Range.FromCol < Dimension.FromCol)
                    {
                        AddParentDeleteCol(i, r);
                        i--;
                    }
                    else if (Quads != null && FitsInQuad(r, out int quadIx))
                    {
                        qadd.Add(quadIx, r);
                        Ranges.RemoveAt(i--);
                    }
                }
            }

            if (Quads != null)
            {
                foreach (var q in Quads)
                {
                    if (q.DeleteCol(fromCol, cols))
                    {
                        ret = true;
                    }
                }
            }

            Ranges.AddRange(l);
            foreach (var q in qadd)
            {
                Quads[q.Key].Add(q.Value);
            }
            return ret;
        }


        private bool FitsInQuad(QuadRangeItem<T> r, out int quadIx)
        {
            for(var i=0;i< Quads.Count;i++)
            {
                if (Quads[i].Dimension.IsInside(r.Range))
                {
                    quadIx = i;
                    return true;
                }
            }
            quadIx = -1;
            return false;
        }

        internal bool InsertRow(int fromRow, int rows, int fromCol, int toCol, out bool overflow)
        {
            overflow = false;
            if (fromRow > Dimension.ToRow)
            {
                return false;
            }

            var l=new List<QuadRangeItem<T>>();
            var ret = false;
            for (int i = 0; i < Ranges.Count; i++)
            {
                var r = Ranges[i];
                if(r.Range.ToRow >= fromRow && (r.Range.ToCol > fromCol && r.Range.FromCol < toCol))
                {
                    if (r.Range.FromCol < fromCol)
                    {
                        l.Add(
                            r.CloneWithNewAddress(r.Range.FromRow, r.Range.FromCol, r.Range.ToRow, fromCol - 1));
                        r.Range.FromCol = fromCol;
                    }
                    if(r.Range.ToCol > toCol)
                    {
                        l.Add(
                            r.CloneWithNewAddress(r.Range.FromRow, toCol+1, r.Range.ToRow, r.Range.ToCol));
                        r.Range.ToCol = toCol;
                    }

                    if (r.Range.InsertRow(fromRow, rows))
                    {
                        if (r.Range.ToRow > Dimension.ToRow)
                        {
                            if(AddParentInsertRow(i, r))
                            {
                                overflow = true;
                            }
                            else
                            {
                                i--;
                                ret = true;
                            }
                        }
                    }
                }
            }
            Ranges.AddRange(l);

            if (Quads != null)
            {
                foreach (var q in Quads)
                {
                    if (q.InsertRow(fromRow, rows, fromCol, toCol, out overflow))
                    {
                        ret = true;
                    }
                }
            }

            return ret;
        }
        internal bool InsertCol(int fromCol, int cols, int fromRow, int toRow, out bool overflow)
        {
            overflow = false;
            if (fromCol > Dimension.ToCol)
            {
                return false;
            }

            var l = new List<QuadRangeItem<T>>();
            var ret = false;
            for (int i = 0; i < Ranges.Count; i++)
            {
                var r = Ranges[i];
                if (r.Range.ToRow >= fromRow && (r.Range.ToRow > fromRow && r.Range.FromRow < toRow))
                {
                    if (r.Range.FromRow < fromRow)
                    {
                        l.Add(
                            r.CloneWithNewAddress(r.Range.FromRow, r.Range.FromCol, fromRow - 1, r.Range.ToCol));
                        r.Range.FromRow = fromRow;
                    }

                    if (r.Range.ToRow > toRow)
                    {
                        l.Add(
                            r.CloneWithNewAddress(toRow + 1, r.Range.FromCol, r.Range.ToRow, r.Range.ToCol));
                        r.Range.ToRow = toRow;
                    }

                    if (r.Range.InsertColumn(fromCol, cols))
                    {
                        if (r.Range.ToCol > Dimension.ToCol)
                        {
                            if (AddParentInsertCol(i, r))
                            {
                                overflow = true;
                            }
                            i--;
                            ret = true;
                        }
                    }
                }
            }
            Ranges.AddRange(l);

            if (Quads != null)
            {
                foreach (var q in Quads)
                {
                    if (q.InsertCol(fromCol, cols, fromRow, toRow, out overflow))
                    {
                        ret = true;
                    }
                }
            }

            return ret;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="i"></param>
        /// <param name="r"></param>
        /// <returns>Returns true if overflow occurres</returns>
        private bool AddParentInsertRow(int i, QuadRangeItem<T> r)
        {
            var p = Parent;
            while (p!=null && p.Dimension.ToRow <= r.Range.ToRow) 
            {
                p = p.Parent;
            }            
            
            if(p==null)
            {
                return true;
            }

            p.Add(r);
            Ranges.RemoveAt(i);
            return false;
        }
        private bool AddParentInsertCol(int i, QuadRangeItem<T> r)
        {
            var p = Parent;
            while (p != null && p.Dimension.ToCol <= r.Range.ToCol)
            {
                p = Parent;
            }

            if (p == null)
            {
                return true;
            }

            p.Add(r);
            Ranges.RemoveAt(i);
            return false;
        }

        private void AddParentDeleteRow(int i, QuadRangeItem<T> r)
        {
            var p = Parent;
            while (p != null && p.Dimension.FromRow > r.Range.FromRow)
            {
                p = Parent;
            }

            if (p == null)
            {
                throw new IndexOutOfRangeException("Quadtree out of range");
            }

            p.Add(r);
            Ranges.RemoveAt(i);
        }
        private void AddParentDeleteCol(int i, QuadRangeItem<T> r)
        {
            var p = Parent;
            while (p != null && p.Dimension.FromCol > r.Range.FromCol)
            {
                p = Parent;
            }

            if (p == null)
            {
                throw new IndexOutOfRangeException("Quadtree out of range");
            }

            p.Add(r);
            Ranges.RemoveAt(i);
        }
        internal void Clear(QuadRange clearedRange)
        {
            if (Quads != null)
            {
                foreach (var q in Quads)
                {
                    if (q.Intersect(clearedRange) != IntersectType.OutSide)
                    {
                        q.Clear(clearedRange);
                    }
                }
            }
            var splitedRanges = new List<QuadRangeItem<T>>();
            for(var i=0;i < Ranges.Count;i++)
            {
                var r=Ranges[i].Range;
                var isct = clearedRange.Intersect(r);
                if (isct == IntersectType.Inside)
                {
                    Ranges.RemoveAt(i--);
                }
                else if(isct == IntersectType.Partial)
                {
                    splitedRanges.AddRange(
                           QuadTreeAddress.ClearAddress(Ranges[i], clearedRange));

                    Ranges.RemoveAt(i--);
                }
            }
            Ranges.AddRange(splitedRanges);
        }
    }
}
