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

        private readonly Dictionary<uint, IndexRelation> _relations = new Dictionary<uint, IndexRelation>();
        private readonly Dictionary<uint, List<uint>> _relationTargets = new Dictionary<uint, List<uint>>();
        private readonly Dictionary<uint, List<uint>> _relationPointers = new Dictionary<uint, List<uint>>();
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

        public void ReIndex()
        {
            foreach(var collection in _collections.Values)
            {
                collection.ReIndex();
            }
        }

        public IndexEndpoint GetItem(uint id)
        {
            foreach(var collection in _collections.Values)
            {
                var result = collection.GetById(id);
                if(result != null)
                {
                    return result;
                }
            }
            return null;
        }

        public int? GetIndexByItem(IndexEndpoint endpoint)
        {
            if (endpoint == null) return null;
            return GetIndexById(endpoint.Id, endpoint.Entity);
        }

        public int? GetIndexById(uint id, RichDataEntities entity)
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

        public bool DeleteRelation(IndexRelation relation)
        {
            if (!_relations.ContainsKey(relation.Id)) return false;
            _relations.Remove(relation.Id);
            if(_relationPointers.ContainsKey(relation.From.Id))
            {
                var pointersList = _relationPointers[relation.From.Id];
                if(pointersList != null && pointersList.Any())
                {
                    pointersList.RemoveAll(x => x == relation.To.Id);
                }
            }
            if(_relationTargets.ContainsKey(relation.To.Id))
            {
                var targetsList = _relationTargets[relation.To.Id];
                if(targetsList != null && targetsList.Any())
                {
                    targetsList.RemoveAll(x => x == relation.From.Id);
                }
            }
            return true;
        }

        public void AddRelation(IndexRelation relation)
        {
            if(!_relationTargets.ContainsKey(relation.To.Id))
            {
                _relationTargets.Add(relation.To.Id, new List<uint>());
            }
            if (!_relationPointers.ContainsKey(relation.From.Id))
            {
                _relationPointers.Add(relation.From.Id, new List<uint>());
            }
            _relationTargets[relation.To.Id].Add(relation.Id);
            _relationPointers[relation.From.Id].Add(relation.Id);
            _relations.Add(relation.Id, relation);
        }

        public void AddRelationWithSubRelations(IndexRelationWithSubRelations relation)
        {
            if (!_relationPointers.ContainsKey(relation.From.Id))
            {
                _relationPointers.Add(relation.From.Id, new List<uint>());
            }
            _relationPointers[relation.From.Id].Add(relation.Id);
            _relations.Add(relation.Id, relation);
        }

        public IEnumerable<IndexRelation> GetRelationTargets(uint id)
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

        public IEnumerable<IndexRelation> GetRelationTargets(uint id, Func<IndexRelation, bool> filter)
        {
            if (_relationPointers.ContainsKey(id))
            {
                var relationIds = _relationPointers[id];
                var result = new List<IndexRelation>();
                foreach (var relId in relationIds)
                {
                    var rel = _relations[relId];
                    if(filter.Invoke(rel))
                    {
                        result.Add(_relations[relId]);
                    }
                }
                return result;
            }
            return Enumerable.Empty<IndexRelation>();
        }

        public IEnumerable<IndexRelation> GetRelationPointers(uint id)
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

        public IEnumerable<IndexRelation> GetRelationPointers(uint id, Func<IndexRelation, bool> filter)
        {
            if (_relationTargets.ContainsKey(id))
            {
                var relationIds = _relationTargets[id];
                var result = new List<IndexRelation>();
                foreach (var relId in relationIds)
                {
                    var rel = _relations[relId];
                    if (filter.Invoke(rel))
                    {
                        result.Add(_relations[relId]);
                    }
                    result.Add(_relations[relId]);
                }
                return result;
            }
            return Enumerable.Empty<IndexRelation>();
        }

        public void EntityDeleted(IndexEndpoint deletedEntity)
        {
            if (deletedEntity == null) return;
            if (!_collections.ContainsKey(deletedEntity.Entity)) return;
            var coll = _collections[deletedEntity.Entity];
            var pointerRels = _relationPointers[deletedEntity.Id];
            if(pointerRels != null && pointerRels.Any())
            {
                foreach(var pointerRelId in pointerRels)
                {
                    var pointerRel = _relations[pointerRelId];
                    var item = GetItem(pointerRel.From.Id);
                    if(item != null)
                    {
                        item.OnConnectedEntityDeleted(deletedEntity.Id, deletedEntity.Entity);
                    }
                }
            }
            var targetsRels = _relationTargets[deletedEntity.Id];
            if(targetsRels != null && targetsRels.Any())
            {
                foreach(var targetRelId in targetsRels)
                {
                    var targetRel = _relations[targetRelId];
                    var item = GetItem(targetRel.From.Id);
                    if (item != null)
                    {
                        item.OnConnectedEntityDeleted(deletedEntity.Id, deletedEntity.Entity);
                    }
                }
            }
        }
    }
}
