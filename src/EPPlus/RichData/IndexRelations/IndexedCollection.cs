/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal abstract class IndexedCollection<T> : IEnumerable<T>, IndexedCollectionInterface
        where T : IndexEndpoint
    {
        public IndexedCollection(RichDataIndexStore store, RichDataEntities entity)
        {
            _list = new List<T>();
            _store = store;
            _entity = entity;
            store.RegisterCollection(entity, this);
        }

        private readonly Dictionary<int, IEnumerable<IndexEndpoint>> _incomingPointers= new Dictionary<int, IEnumerable<IndexEndpoint>>();
        private readonly Dictionary<int, IEnumerable<IndexEndpoint>> _outgoingPointers = new Dictionary<int, IEnumerable<IndexEndpoint>>();
        private readonly Dictionary<uint, int> _idToIndex = new Dictionary<uint, int>();
        private readonly Dictionary<uint, T> _items = new Dictionary<uint, T>();
        private readonly List<T> _list;
        private readonly RichDataIndexStore _store;
        private readonly RichDataEntities _entity;

        /// <summary>
        /// Returns Id:s of all instances of other entities that points to this record
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        protected IEnumerable<IndexEndpoint> GetIncomingPointers(int id)
        {
            if(_incomingPointers.ContainsKey(id))
            {
                return _incomingPointers[id];
            }
            else
            {
                return Enumerable.Empty<IndexEndpoint>();
            }
        }

        public void ReIndex()
        {
            var ix = 0;
            foreach(IndexEndpoint item in this)
            {
                if(item.Deleted)
                {
                    _idToIndex.Remove(item.Id);
                }
                else
                {
                    _idToIndex[item.Id] = ix++;
                }
            }
        }

        /// <summary>
        /// Returns Id:s of all instances of other entities that this record points to
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        protected IEnumerable<IndexEndpoint> GetOutgoingPointers(int id)
        {
            if (_outgoingPointers.ContainsKey(id))
            {
                return _outgoingPointers[id];
            }
            else
            {
                return Enumerable.Empty<IndexEndpoint>();
            }
        }

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

        public virtual RichDataEntities EntityType => _entity;

        int IndexedCollectionInterface.Count => _list.Count;

        public Type IndexedType => typeof(T);

        IndexEndpoint IndexedCollectionInterface.this[int index] => this[index];

        public virtual void Add(T item)
        {
            _idToIndex.Add(item.Id, _list.Count);
            _items.Add(item.Id, item);
            _list.Add(item);
        }

        public virtual T Get(uint id)
        {
            if (!_idToIndex.ContainsKey(id)) return null;
            var ix = _idToIndex[id];
            return _list[ix];
        }

        public virtual bool Remove(T item)
        {
            if(item != null && _idToIndex.ContainsKey(item.Id))
            {
                _idToIndex.Remove(item.Id);
            }
            return _list.Remove(item);
        }

        public virtual void RemoveAt(int index)
        {
            var item = _list[index];
            if(item != null)
            {
                _idToIndex.Remove(item.Id);
                _list.RemoveAt(index);
            }
        }

        protected T GetItemById(uint id)
        {
            if (!_idToIndex.ContainsKey(id)) return default;
            var ix = _idToIndex[id];
            return _list[ix];
        }

        public void DeleteEndpoint(uint id)
        {
            var endpoint = _list.FirstOrDefault(x => x.Id == id);
            if(endpoint != null)
            {
                var ix = _list.IndexOf(endpoint);
                //TODO: re-index?
            }
        }

        public IndexRelation CreateRelation(IndexEndpoint from, IndexEndpoint to, IndexType indexType)
        {
            return _store.CreateAndAddRelation(from, to, indexType);
        }

        public IndexRelation CreateRelation(IndexEndpoint from, int toIndex, IndexType indexType)
        {
            var to = this[toIndex];
            return _store.CreateAndAddRelation(from, to, indexType);
        }

        public T GetItem(uint id)
        {
            if (!_items.ContainsKey(id)) return null;
            return _items[id];
        }

        public int GetNextIndex()
        {
            return _list.Count;
        }

        public int FindIndex(Predicate<T> match)
        {
            return _list.FindIndex(match);
        }

        public int? GetIndexById(uint id)
        {
            if (!_idToIndex.ContainsKey(id)) return default;
            return _idToIndex[id];
        }

        int? IndexedCollectionInterface.GetIndexById(uint id)
        {
            return GetIndexById(id);
        }

        IndexEndpoint IndexedCollectionInterface.GetById(uint id)
        {
            if(_items.ContainsKey(id)) return _items[id];
            return null;
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
