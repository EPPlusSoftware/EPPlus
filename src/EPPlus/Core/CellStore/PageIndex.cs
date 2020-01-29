/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;

namespace OfficeOpenXml.Core.CellStore
{
    internal class PageIndex : IndexBase, IDisposable
    {
        public PageIndex(int pageSizeMin)
        {
            Rows = new IndexItem[pageSizeMin];
            RowCount = 0;
        }
        public PageIndex(IndexItem[] rows, int count)
        {
            Rows = rows;
            RowCount = count;
        }
        public PageIndex(PageIndex pageItem, int start, int size)
            : this(pageItem, start, size, pageItem.Index, pageItem.Offset)
        {

        }
        public PageIndex(PageIndex pageItem, int start, int size, short index, int offset)
        {
            Rows = new IndexItem[CellStore<int>.GetSize(size)];
            Array.Copy(pageItem.Rows, start, Rows, 0, pageItem.RowCount-start);
            RowCount = size;    
            Index = index;
            Offset = offset;
        }
        ~PageIndex()
        {
            Rows = null;
        }
        internal int Offset = 0;
        /// <summary>
        /// Rows in the rows collection. 
        /// </summary>
        internal int RowCount;
        internal int IndexOffset
        {
            get
            {
                return IndexExpanded + Offset;
            }
        }
        internal int IndexExpanded
        {
            get
            {
                return (Index << CellStoreSettings._pageBits);
            }
        }
        internal IndexItem[] Rows { get; set; }
        /// <summary>
        /// First row index minus last row index
        /// </summary>
        internal int RowSpan
        {
            get
            {
                return MaxIndex - MinIndex+1;
            }
        }

        internal int GetPosition(int offset)
        {
            return ArrayUtil.OptimizedBinarySearch(Rows, offset, RowCount);
        }
        internal int GetRowPosition(int row)
        {
            var offset = row - IndexOffset;
            return ArrayUtil.OptimizedBinarySearch(Rows, offset, RowCount);
        }
        internal int GetNextRow(int row)
        {
            var o = GetRowPosition(row);
            if (o < 0)
            {
                o = ~o;
                if (o < RowCount)
                {
                    return o;
                }
                else
                {
                    return -1;
                }
            }
            return o;
        }

        public int MinIndex
        {
            get
            {
                if (RowCount > 0)
                {
                    return IndexOffset + Rows[0].Index;
                }
                else
                {
                    return -1;
                }
            }
        }
        public int MaxIndex
        {
            get
            {
                if (RowCount > 0)
                {
                    return IndexOffset + Rows[RowCount - 1].Index;
                }
                else
                {
                    return -1;
                }
            }
        }
        public int GetIndex(int pos)
        {
            return IndexOffset + Rows[pos].Index;
        }
        public void Dispose()
        {
            Rows = null;
        }

        internal bool IsWithin(int fromRow, int toRow)
        {
            return fromRow <= MinIndex  && toRow >= MaxIndex;
        }
        internal bool StartsWithin(int fromRow, int toRow)
        {
            return fromRow <= MaxIndex && toRow >= MinIndex;
        }

        internal bool StartsAfter(int row)
        {
            return row > MaxIndex;
        }

        internal int GetRow(int rowIx)
        {
            return IndexOffset + Rows[rowIx].Index;
        }
    }
}