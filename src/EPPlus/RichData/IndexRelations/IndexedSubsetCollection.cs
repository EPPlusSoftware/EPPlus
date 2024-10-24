using OfficeOpenXml.RichData.IndexRelations.EventArguments;
using System;
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
        public EventHandler<CollectionIsEmptyEventArgs> CollectionIsEmpty;

        public void Add(T item)
        {
            item.EndpointDeleted += OnEndpointDeleted;
            //TODO: should we handle empty collection via events?
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

        private void OnEndpointDeleted(object source, EndpointDeletedEventArgs e)
        {
            _items.Remove(e.Id);
            if(_items.Count == 0)
            {
                var e2 = new CollectionIsEmptyEventArgs(e.Deletions);
                CollectionIsEmpty?.Invoke(this, e2);
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
