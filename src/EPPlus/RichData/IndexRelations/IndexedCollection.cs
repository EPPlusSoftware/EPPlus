using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal class IndexedCollection<T> : IEnumerable<T>
        where T : IndexedValue
    {
        public IndexedCollection()
        {
            _list = new List<T>();
        }

        private readonly Dictionary<int, List<IndexPointer>> _pointers= new Dictionary<int, List<IndexPointer>>();
        private readonly List<T> _list;
        public IEnumerator<T> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        public virtual int Count
        {
            get
            {
                return _list.Count;
            }
        }

        public virtual void Add(T item)
        {
            _list.Add(item);
        }

        public virtual bool Remove(T item)
        {
            return _list.Remove(item);
        }

        public virtual void RemoveAt(int index)
        {
            _list.RemoveAt(index);
        }

        public T this[int index]
        {
            get
            {
                return _list[index];
            }
            set
            {
                _list[index] = value;
            }
        }
    }
}
