using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Core
{
    public class EPPlusReadOnlyList<T> : IEnumerable<T>
    {
        internal List<T> _list=new List<T>();
        public IEnumerator GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator<T> IEnumerable<T>.GetEnumerator()
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
        public T this[T value]
        {
            get
            {
                var ix=_list.IndexOf(value);
                if(ix<0)
                {
                    return _list[ix];
                }
                return default;
            }
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
