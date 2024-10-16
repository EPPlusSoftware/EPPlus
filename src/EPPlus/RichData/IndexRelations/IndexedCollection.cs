﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal abstract class IndexedCollection<T> : IEnumerable<T>, IndexedCollectionInterface
        where T : IndexEndpoint
    {
        public IndexedCollection(ExcelRichData richData, RichDataEntities entity)
        {
            _list = new List<T>();
            _richData = richData;
            richData.IndexStore.RegisterCollection(entity, this);
        }

        private readonly Dictionary<int, IEnumerable<int>> _incomingPointers= new Dictionary<int, IEnumerable<int>>();
        private readonly Dictionary<int, IEnumerable<int>> _outgoingPointers = new Dictionary<int, IEnumerable<int>>();
        private readonly Dictionary<int, int> _idToIndex = new Dictionary<int, int>();
        private readonly List<T> _list;
        private readonly ExcelRichData _richData;

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

        IndexEndpoint IndexedCollectionInterface.this[int index] => this[index];

        public virtual void Add(T item)
        {
            _idToIndex.Add(item.Id, _list.Count);
            _list.Add(item);
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
            _richData.IndexStore.AddRelation(relation);
            return relation;
        }

        public IndexRelation CreateRelation(IndexEndpoint from, int toIndex, IndexType indexType)
        {
            var to = this[toIndex];
            var relation = new IndexRelation(from, to, indexType);
            _richData.IndexStore.AddRelation(relation);
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
