using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal class RichDataIndexStore
    {
        public RichDataIndexStore()
        {
            
        }

        private readonly Dictionary<int, IndexRelation> _relations = new Dictionary<int, IndexRelation>();
        private readonly Dictionary<int, List<int>> _relationTargets = new Dictionary<int, List<int>>();
        private readonly Dictionary<int, List<int>> _relationPointers = new Dictionary<int, List<int>>();
        private readonly Dictionary<RichDataEntities, IndexedCollectionInterface> _collections = new Dictionary<RichDataEntities, IndexedCollectionInterface>();
        private readonly Dictionary<Type, RichDataEntities> _typeToEntity = new Dictionary<Type, RichDataEntities>();

        public void RegisterCollection(RichDataEntities entity, IndexedCollectionInterface coll)
        {
            if(!_collections.ContainsKey(entity))
            {
                _collections.Add(entity, coll);
            }
            var t = coll.IndexedType;
            if(!_typeToEntity.ContainsKey(t))
            {
                _typeToEntity[t] = entity;
            }
        }

        public RichDataEntities? GetEntityByType(Type t)
        {
            if(_typeToEntity.ContainsKey(t))
            {
                return _typeToEntity[t];
            }
            return default;
        }

        public int? GetIndexById(int id, RichDataEntities entity)
        {
            var coll = _collections[entity];
            return coll.GetIndexById(id);
        }

        public int GetNextIndex(RichDataEntities entity)
        {
            if(!_collections.ContainsKey(entity))
            {
                throw new ArgumentException("No collection registred for entity " + entity.ToString());
            }
            var coll = _collections[entity];
            return coll.GetNextIndex();
        }

        public void AddRelation(IndexRelation relation)
        {
            if(!_relationTargets.ContainsKey(relation.To.Id))
            {
                _relationTargets.Add(relation.To.Id, new List<int>());
            }
            if (!_relationPointers.ContainsKey(relation.From.Id))
            {
                _relationPointers.Add(relation.From.Id, new List<int>());
            }
            _relationTargets[relation.To.Id].Add(relation.Id);
            _relationPointers[relation.From.Id].Add(relation.Id);
            _relations.Add(relation.Id, relation);
        }

        public IEnumerable<IndexRelation> GetRelationTargets(int id)
        {
            if(_relationPointers.ContainsKey(id))
            {
                var relationIds = _relationPointers[id];
                var result = new List<IndexRelation>();
                foreach(var relId in relationIds)
                {
                    result.Add(_relations[relId]);
                }
                return result;
            }
            return Enumerable.Empty<IndexRelation>();
        }

        public IEnumerable<IndexRelation> GetRelationPointers(int id)
        {
            if (_relationTargets.ContainsKey(id))
            {
                var relationIds = _relationTargets[id];
                var result = new List<IndexRelation>();
                foreach (var relId in relationIds)
                {
                    result.Add(_relations[relId]);
                }
                return result;
            }
            return Enumerable.Empty<IndexRelation>();
        }
    }
}
