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
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Core
{
    internal class ChangeableDictionary<T> : IEnumerable<T>
    {
        internal int[][] _index;
        internal List<T> _items;
        internal int _count;
        int _defaultSize;
        internal ChangeableDictionary(int size = 8)
        {
            _defaultSize = size;
            Clear();
        }

        internal T this[int key]
        {
            get
            {
                var pos = Array.BinarySearch(_index[0], 0, _count, key);
                if(pos>=0)
                {
                    return _items[_index[1][pos]];
                }
                else
                {
                    return default(T);
                }
            }
        }

        internal void InsertAndShift(int fromPosition, int add)
        {
            var pos = Array.BinarySearch(_index[0], 0, _count, fromPosition);
            if(pos<0)
            {
                pos = ~pos;
            }
            Array.Copy(_index[0], pos, _index[0], pos + 1, _count - pos);
            Array.Copy(_index[1], pos, _index[1], pos + 1, _count - pos);
            _count++;
            for (int i=pos;i<Count;i++)
            {
                _index[0][i] += add;
            }
        }

        internal int Count { get { return _count; } }

        public void Add(int key, T value)
        {
            var pos = Array.BinarySearch(_index[0], 0, _count, key);
            if (pos >= 0)
            {
                throw (new ArgumentException("Key already exists"));
            }
            pos = ~pos;
            if (pos >= _index[0].Length)
            {
                Array.Resize(ref _index[0], _index[0].Length << 1);
                Array.Resize(ref _index[1], _index[1].Length << 1);
            }
            if (pos < Count)
            {
                Array.Copy(_index[0], pos, _index[0], pos + 1, _index[0].Length - pos - 1);
                Array.Copy(_index[1], pos, _index[1], pos + 1, _index[1].Length - pos - 1);
            }
            _count++;
            _index[0][pos] = key;
            _index[1][pos] = _items.Count;
            _items.Add(value);
        }

        internal void Move(int fromPosition, int toPosition, bool before)
        {
            if (Count <= 1 || fromPosition == toPosition) return;
            var listItem = _index[1][fromPosition];
            var insertPos = before ? toPosition : toPosition + 1;
            var removePos = fromPosition;
            if(insertPos>removePos)
            {
                InsertAndShift(insertPos, 1);
                RemoveAndShift(removePos, false);
                insertPos--;
            }
            else
            {
                RemoveAndShift(removePos, false);
                InsertAndShift(insertPos, 1);
            }
            _index[0][insertPos] = insertPos;
            _index[1][insertPos] = listItem;
        }

        public void Clear()
        {
            _index = new int[2][];
            _index[0] = new int[_defaultSize];
            _index[1] = new int[_defaultSize];
            _items = new List<T>();
        }

        public bool ContainsKey(int key)
        {
            return Array.BinarySearch(_index[0], 0, _count, key) >= 0;
        }
    
        public IEnumerator<T> GetEnumerator()
        {
            return new ChangeableDictionaryEnumerator<T>(this);
        }

        public bool RemoveAndShift(int key)
        {
            return RemoveAndShift(key, true);
        }

        private bool RemoveAndShift(int key, bool dispose)
        {
            var pos = Array.BinarySearch(_index[0], 0, _count, key);
            if (pos >= 0)
            {
                if (dispose)
                {
                    (_items[_index[1][pos]] as IDisposable)?.Dispose();
                    _items[_index[1][pos]] = default(T);
                }

                if (pos < Count)
                {
                    Array.Copy(_index[0], pos + 1, _index[0], pos, Count - pos - 1);
                    Array.Copy(_index[1], pos + 1, _index[1], pos, Count - pos - 1);
                }
                _count--;
                for (var i = pos; i < _count; i++)
                {
                    _index[0][i]--;
                }
                return true;
            }
            return false;
        }

        public bool TryGetValue(int key, out T value)
        {
            var pos = Array.BinarySearch(_index[0], 0, _count, key);
            if (pos >= 0)
            {
                value = _items[pos];
                return true;
            }
            else
            {
                value = default(T);
                return false;
            }
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return new ChangeableDictionaryEnumerator<T>(this);
        }
    }
    internal class ChangeableDictionaryEnumerator<T> : IEnumerator<T>
    {
        int _index=-1;
        ChangeableDictionary<T> _ts;
        public ChangeableDictionaryEnumerator(ChangeableDictionary<T> ts)
        {
            _ts = ts;
        }
        public T Current
        {
            get
            {
                if (_index >= _ts._count)
                {
                    return default(T);
                }
                else
                {
                    return _ts._items[_ts._index[1][_index]];
                }
            }
        }

        object IEnumerator.Current => Current;

        public void Dispose()
        {
            _ts = null;
        }

        public bool MoveNext()
        {
            _index++;
            if (_ts.Count == _index)
            {
                return false;
            }
            return true;
        }

        public void Reset()
        {
            throw new NotImplementedException();
        }
    }
}
