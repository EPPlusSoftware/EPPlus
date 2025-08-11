using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.Worksheet.XmlWriter;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core
{
    internal class ColumnSpanDictionary<T>: ChangeableDictionary<T> where T : ExcelColumn
    {
        int largestCol = -1;
        int smallestCol = -1;

        //public override void Add(int key, T value)
        //{
        //    //largestCol = largestCol < value.ColumnMax ? largestCol : value.ColumnMax;
        //    //smallestCol = smallestCol < value.ColumnMin ? smallestCol : value.ColumnMin;

        //    //base.Add(key, value);
        //    //var min = value.ColumnMin;
        //    //var max = value.ColumnMax;
        //    //var pos = Array.BinarySearch(_index[0], 0, _count, key);
        //    //if (pos >= 0)
        //    //{
        //    //    throw (new ArgumentException("Key already exists"));
        //    //}
        //    //pos = ~pos;
        //    //if (pos >= _index[0].Length - 1)
        //    //{
        //    //    Array.Resize(ref _index[0], _index[0].Length << 1);
        //    //    Array.Resize(ref _index[1], _index[1].Length << 1);
        //    //}
        //    //if (pos < Count)
        //    //{
        //    //    Array.Copy(_index[0], pos, _index[0], pos + 1, _index[0].Length - pos - 1);
        //    //    Array.Copy(_index[1], pos, _index[1], pos + 1, _index[1].Length - pos - 1);
        //    //}
        //    //_count++;


        //    //_index[0][pos] = key;
        //    //_index[1][pos] = _items.Count;
        //    //_items.Add(value);
        //    //Version++;
        //}
        //public override bool RemoveAndShift(int key)
        //{
        //    //var removed = base.RemoveAndShift(key);
        //    //if(removed)
        //    //{
        //    //    var lastCol = _items.Last();
        //    //    _items.Colu
        //    //}
        //    //return removed;
        //}

        internal int FindInternalIndex(int col)
        {
            int closestIndex = -1;

            if(_items.Count <= 0)
            {
                return closestIndex;
            }

            var columnIndex = Array.BinarySearch(_index[0], 0, _count, col);

            var indexExists = columnIndex >= 0;
            var indexNotFound = columnIndex < 0;

            if (indexNotFound)
            {
                var inverted = ~columnIndex;

                bool indexBetweenTwoColMin = inverted < _items.Count && inverted > 0;
                bool indexLargerThanLargestColMin = inverted >= _items.Count;

                if (indexBetweenTwoColMin || indexLargerThanLargestColMin)
                {
                    /*
                     * In Array.Binary search:
                     * "If value is not found and value is less than one or more elements in array;
                     * the negative number returned is the bitwise complement of the index of the first element that is larger than value"
                     * And since we store based on colMin. It will find the index above the one we want.
                     * 
                     * Similarily even if we are larger than the largest colMin the colMax could still be larger
                     */
                    inverted -= 1;
                }
           
                if (inverted < _index[1].Length - 1)
                {
                    closestIndex = _index[1][inverted];
                }
                else
                {
                    //if inverted is less than collection maximum it does not exist. index is already -1.
                }
            }
            else
            {
                closestIndex = _index[1][columnIndex];
            }
            return closestIndex;
        }

        internal ExcelColumn GetExcelColumn(int col)
        {
            var internalIndex = FindInternalIndex(col);
            if (internalIndex > -1 && _items.Count > 0)
            {
                var closestCol = _items[internalIndex];
                if(closestCol != null)
                {
                    if (closestCol.ColumnMin <= col && col <= closestCol.ColumnMax)
                    {
                        return closestCol;
                    }
                }
            }
            return null;
        }

        internal bool TryGetExcelColumn(int col, out ExcelColumn columnValue)
        {
            columnValue = GetExcelColumn(col);
            return columnValue != null;
        }

        internal void UpdateColumn(int col, ExcelColumn columnValue)
        {
            var pos = Array.BinarySearch(_index[0], 0, _count, col);
            if (pos < 0)
            {
                Add(columnValue.ColumnMin, (T)columnValue);
            }
            //if(!TryGetExcelColumn(col, out ExcelColumn existingCol))
            //{
            //    Add(columnValue.ColumnMin, (T)columnValue);
            //    //this[col] = (T)columnValue;
            //    //Add(columnValue.ColumnMin, (T)columnValue);
            //}
            //else
            //{
            //    Add(columnValue.ColumnMin, (T)columnValue);
            //}
        }

        internal void UpdateDeletedPositions(int columnFrom, int columnTo)
        {
            //List<int> indexesToDelete = new List<int>();
            for (int i = columnTo; i > columnFrom; i--)
            {
                var index = FindInternalIndex(i);
                if(index > -1)
                {
                    RemoveAndShift(index, false);
                }
            }
        }
        ////ExcelColumn is updated via GetValueInner changes but the positions remain unchanged.
        ////Ensure positions have changed if ColumnMin has
        //internal void UpdateDictPositions(int oldMin, int newMin)
        //{
        //    var index = FindInternalIndex(oldMin);
        //    if (index != -1)
        //    {
        //        if (newMin < oldMin)
        //        {
        //            InsertAndShift(newMin, newMin - oldMin);
        //            RemoveAndShift(oldMin, false);
        //            index--;
        //        }
        //        else
        //        {
        //            RemoveAndShift(oldMin, false);
        //            InsertAndShift(newMin, newMin - oldMin);
        //        }
        //        _index[0][index] = newMin;
        //        Version++;

        //        //Move(index,)
        //        //_index[0][index] = newMin;
        //        //Move(index,newMin, oldMin < newMin);
        //        //if (newMin < oldMin)
        //        //{
        //        //    InsertAndShift(newMin, newMin - oldMin);
        //        //    RemoveAndShift(oldMin,false);
        //        //    //insertPos--;
        //        //}
        //        //else
        //        //{
        //        //    RemoveAndShift(oldMin, false);
        //        //    InsertAndShift(newMin, newMin - oldMin);
        //        //}
        //        //////Move(oldMin,newMin,)
        //        ////RemoveAndShift(oldMin);
        //        ////InsertAndShift(newMin, newMin - oldMin);
        //        //if(Count > _index[0].Length -1)
        //        //{
        //        //    Array.Resize(ref _index[0], _index[0].Length << 1);
        //        //    Array.Resize(ref _index[1], _index[1].Length << 1);
        //        //}
        //        //if(Count >= _index[0].Length - 1)
        //        //{

        //        //}
        //    }
        //    //var colExists = TryGetExcelColumn(oldMin, out ExcelColumn existingCol);
        //    //if (colExists)
        //    //{
        //    //    FindInternalIndex(oldMin);
        //    //}
        //}

        //public override bool RemoveAndShift(int key)
        //{
        //    var colMax = this[key].ColumnMax;
        //    if (this[key].ColumnMax == largestCol)
        //    {

        //    }
        //    var removed = base.RemoveAndShift(key);
        //    return removed;
        //}

        //internal bool HasColumn(int col)
        //{
        //    if ( smallestCol <= col && col <= largestCol)
        //    {
        //        return true;
        //    }
        //    else
        //    {
        //        return false;
        //    }
        //}


        //internal int OptimizedBinarySearch(int col, int wsColMax, int wsColMin)
        //{
        //    var store = _items;

        //    if (wsColMax == 0) return -1;
        //    int low = wsColMin, high = wsColMax - 1, mid;

        //    while (low <= high)
        //    {
        //        mid = (low + high) >> 1;

        //        if (col < store[mid].ColumnMax)
        //            high = mid - 1;

        //        else if (col > store[mid].ColumnMin)
        //            low = mid + 1;

        //        else
        //            return mid;
        //    }
        //    return ~low;
        //}

        //internal ExcelColumn GetExcelColumn(int col)
        //{

        //    if(largestCol > -1 && smallestCol > -1)
        //    {
        //        var index = OptimizedBinarySearch(col, largestCol, smallestCol);
        //        return _items[col];
        //    }
        //    //OptimizedBinarySearch(col, largestCol, smallestCol);
        //    ////var colCompare = new ExcelColumnComparer()
        //    ////    _items.BinarySearch(0,_items.)
        //    ////Array.BinarySearch(_items, col, colCompare);

        //    //foreach (var item in _items)
        //    //{
        //    //    if (item.ColumnMin <= col && col <= item.ColumnMax)
        //    //    {
        //    //        return item;
        //    //    }
        //    //}
        //    return null;
        //    //ArrayUtil.OptimizedBinarySearch(_items, col, _count);
        //    //if(HasColumn(col))
        //    //{

        //    //}
        //    //return null;
        //}
    }
}
