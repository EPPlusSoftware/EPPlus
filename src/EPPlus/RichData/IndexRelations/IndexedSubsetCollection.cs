using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal class IndexedSubsetCollection<T>
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

        public T First()
        {
            return FirstOrDefault();
        }

        public T First(Predicate<T> predicate)
        {
            return FirstOrDefault(predicate);
        }

        public T FirstOrDefault()
        {
            if (_items.Count == 0)
            {
                return null;
            }
            var id = _items[0];
            return _collection.GetItem(id) as T;
        }

        public T FirstOrDefault(Predicate<T> predicate)
        {
            if (_items.Count == 0)
            {
                return null;
            }
            for (var ix = 0; ix < _items.Count; ix++)
            {
                var id = _items[ix];
                var item = _collection.GetItem(id) as T;
                if (item != default && predicate.Invoke(item))
                {
                    return item;
                }
            }
            return null;
        }

        public T Get(uint id)
        {
            if(!_itemIds.Contains(id))
            {
                return null;
            }
            return _collection.Get(id);
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
