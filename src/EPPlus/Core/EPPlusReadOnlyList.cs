using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Core
{
    /// <summary>
    /// A readonly collection of a generic type
    /// </summary>
    /// <typeparam name="T">The generic type</typeparam>
    public class EPPlusReadOnlyList<T> : IEnumerable<T>
    {
        internal List<T> _list=new List<T>();
        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        public IEnumerator GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        public T this[int index]
        {
            get
            {
                return _list[index];
            }
        }
        public T GetByValue(T value)
        {
            var ix=_list.IndexOf(value);
            if(ix<0)
            {
                return _list[ix];
            }
            return default;
        }
        public int GetIndexByValue(T value)
        {
            return _list.IndexOf(value);
        }
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
        internal void Clear()
        {
            _list.Clear();
        }
        internal void Add(T item)
        {
            _list.Add(item);
        }

    }
}
