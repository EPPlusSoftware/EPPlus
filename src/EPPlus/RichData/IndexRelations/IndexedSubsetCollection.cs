﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    /// <summary>
    /// A filter on an <see cref="IndexedCollection{T}"/>. Only values that will are added via this wrapper will be returned.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal class IndexedSubsetCollection<T> : IEnumerable<T>
        where T : IndexEndpoint
    {
        public IndexedSubsetCollection(IndexedCollection<T> coll)
        {
            _collection = coll;
        }

        private readonly IndexedCollection<T> _collection;
        private readonly HashSet<uint> _itemIds = new HashSet<uint>();
        private readonly List<uint> _items = new List<uint>();

        public void Add(T item)
        {
            if (_collection.GetItem(item.Id) == null)
            {
                _collection.Add(item);
            }
            if(!_itemIds.Contains(item.Id))
            {
                _itemIds.Add(item.Id);
                _items.Add(item.Id);
            }
        }

        public int Count => _items.Count;

        public T Get(uint id)
        {
            if(!_itemIds.Contains(id))
            {
                return null;
            }
            return _collection.Get(id);
        }

        public IEnumerator<T> GetEnumerator()
        {
            foreach(var id in _items)
            {
                yield return _collection.Get(id);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public T this[int index]
        {
            get
            {
                var id = _items[index];
                return _collection.GetItem(id);
            }
        }
    }
}
