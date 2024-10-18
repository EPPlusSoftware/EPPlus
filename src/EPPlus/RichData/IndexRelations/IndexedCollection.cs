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
            store.RegisterCollection(entity, this);
        }

        private readonly Dictionary<int, IEnumerable<int>> _incomingPointers= new Dictionary<int, IEnumerable<int>>();
        private readonly Dictionary<int, IEnumerable<int>> _outgoingPointers = new Dictionary<int, IEnumerable<int>>();
        private readonly Dictionary<int, int> _idToIndex = new Dictionary<int, int>();
        private readonly List<T> _list;
        private readonly RichDataIndexStore _store;

        /// <summary>
        /// Returns Id:s of all instances of other entities that points to this record
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        protected IEnumerable<int> GetIncomingPointers(int id)
        {
            if(_incomingPointers.ContainsKey(id))
            {
                return _incomingPointers[id];
            }
            else
            {
                return Enumerable.Empty<int>();
            }
        }

        /// <summary>
        /// Returns Id:s of all instances of other entities that this record points to
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        protected IEnumerable<int> GetOutgoingPointers(int id)
        {
            if (_outgoingPointers.ContainsKey(id))
            {
                return _outgoingPointers[id];
            }
            else
            {
                return Enumerable.Empty<int>();
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

        public abstract RichDataEntities EntityType { get; }

        RichDataEntities IndexedCollectionInterface.EntityType => EntityType;

        int IndexedCollectionInterface.Count => _list.Count;

        public Type IndexedType => typeof(T);

        IndexEndpoint IndexedCollectionInterface.this[int index] => this[index];

        public virtual void Add(T item)
        {
            _idToIndex.Add(item.Id, _list.Count);
            _list.Add(item);
        }

        public virtual T Get(int id)
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

        protected T GetItemById(int id)
        {
            if (!_idToIndex.ContainsKey(id)) return default;
            var ix = _idToIndex[id];
            return _list[ix];
        }

        public void DeleteEndpoint(int id)
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
            var relation = new IndexRelation(from, to, indexType);
            _store.AddRelation(relation);
            return relation;
        }

        public IndexRelation CreateRelation(IndexEndpoint from, int toIndex, IndexType indexType)
        {
            var to = this[toIndex];
            var relation = new IndexRelation(from, to, indexType);
            _store.AddRelation(relation);
            return relation;
        }

        public T GetItem(int id)
        {
            var ix = _idToIndex[id];
            return _list[ix];
        }

        public int GetNextIndex()
        {
            return _list.Count;
        }

        public int FindIndex(Predicate<T> match)
        {
            return _list.FindIndex(match);
        }

        public int? GetIndexById(int id)
        {
            if (!_idToIndex.ContainsKey(id)) return default;
            return _idToIndex[id];
        }

        int? IndexedCollectionInterface.GetIndexById(int id)
        {
            return GetIndexById(id);
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
