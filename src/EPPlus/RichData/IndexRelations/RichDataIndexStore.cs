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
        private readonly Dictionary<uint, List<IndexRelation>> _outgoingRelations = new Dictionary<uint, List<IndexRelation>>();
        private readonly Dictionary<uint, List<IndexRelation>> _incomingRelations = new Dictionary<uint, List<IndexRelation>>();
        private readonly Dictionary<RichDataEntities, IndexedCollectionInterface> _collections = new Dictionary<RichDataEntities, IndexedCollectionInterface>();
        private readonly Dictionary<Type, RichDataEntities> _typeToEntity = new Dictionary<Type, RichDataEntities>();
        private readonly IdGenerator _idGenerator = new IdGenerator();

        public uint GetNewId()
        {
            return _idGenerator.GetNewId();
        }

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
            return GetIndexById(endpoint.Id, endpoint.EntityType);
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
            if(_incomingRelations.ContainsKey(relation.To.Id))
            {
                var pointersList = _incomingRelations[relation.To.Id];
                if(pointersList != null && pointersList.Any())
                {
                    pointersList.RemoveAll(x => x.To == relation.To);
                }
            }
            if(_outgoingRelations.ContainsKey(relation.From.Id))
            {
                var targetsList = _outgoingRelations[relation.From.Id];
                if(targetsList != null && targetsList.Any())
                {
                    targetsList.RemoveAll(x => x.From == relation.From);
                }
            }
            return true;
        }

        public IndexRelation CreateAndAddRelation(IndexEndpoint from, IndexEndpoint to, IndexType indexType)
        {
            var rel = new IndexRelation(this, from, to, indexType);
            AddRelation(rel);
            return rel;
        }

        public IndexRelationWithSubRelations CreateAndAddRelationWithSubRelations(IndexEndpoint endPoint, RichDataEntities entity)
        {
            var rel = new IndexRelationWithSubRelations(this, endPoint, RichDataEntities.MetadataType, IndexType.SubRelations);
            AddRelationWithSubRelations(rel);
            return rel;
        }

        public void AddRelation(IndexRelation relation)
        {
            if(!_outgoingRelations.ContainsKey(relation.From.Id))
            {
                _outgoingRelations.Add(relation.From.Id, new List<IndexRelation>());
            }
            if (!_incomingRelations.ContainsKey(relation.To.Id))
            {
                _incomingRelations.Add(relation.To.Id, new List<IndexRelation>());
            }
            _outgoingRelations[relation.From.Id].Add(relation);
            _incomingRelations[relation.To.Id].Add(relation);
            _relations.Add(relation.Id, relation);
        }

        public void AddRelationWithSubRelations(IndexRelationWithSubRelations relation)
        {
            if (!_incomingRelations.ContainsKey(relation.To.Id))
            {
                _incomingRelations.Add(relation.To.Id, new List<IndexRelation>());
            }
            _incomingRelations[relation.To.Id].Add(relation);
            _relations.Add(relation.Id, relation);
        }

        public IEnumerable<IndexRelation> GetOutgoingRelations(uint id)
        {
            if (_outgoingRelations.ContainsKey(id))
            {
                var relations = _outgoingRelations[id];
                if (relations == null) return Enumerable.Empty<IndexRelation>();
                return relations;
            }
            return Enumerable.Empty<IndexRelation>();
        }

        public IEnumerable<IndexRelation> GetOutgoingRelations(uint id, Func<IndexRelation, bool> filter)
        {
            if (_outgoingRelations.ContainsKey(id))
            {
                var relations = _outgoingRelations[id];
                var result = new List<IndexRelation>();
                foreach (var rel in relations)
                {
                    if (filter.Invoke(rel))
                    {
                        result.Add(rel);
                    }
                }
                return result;
            }
            return Enumerable.Empty<IndexRelation>();
        }

        public IEnumerable<IndexRelation> GetIncomingRelations(uint id)
        {
            if (_incomingRelations.ContainsKey(id))
            {
                var relations = _incomingRelations[id];
                if(relations == null) return Enumerable.Empty<IndexRelation>();
                return relations;
            }
            return Enumerable.Empty<IndexRelation>();
        }

        public IEnumerable<IndexRelation> GetIncomingRelations(uint id, Func<IndexRelation, bool> filter)
        {
            if (_incomingRelations.ContainsKey(id))
            {
                var relations = _incomingRelations[id];
                var result = new List<IndexRelation>();
                foreach (var rel in relations)
                {
                    if (filter.Invoke(rel))
                    {
                        result.Add(rel);
                    }
                }
                return result;
            }
            return Enumerable.Empty<IndexRelation>();
        }

        public void AddSubRelation(uint parentRelId, IndexRelation subRel)
        {
            if(_relations.ContainsKey(parentRelId))
            {
                var parentRel = _relations[parentRelId].AsRelationWithSubRelations();
                if(parentRel != null)
                {
                    subRel.Parent = parentRel;
                    parentRel.SubRelations.Add(subRel);
                }
            }
        }

        public void EntityDeleted(IndexEndpoint deletedEntity)
        {
            if (deletedEntity == null) return;
            if (!_collections.ContainsKey(deletedEntity.EntityType)) return;
            var coll = _collections[deletedEntity.EntityType];
            if (_incomingRelations.ContainsKey(deletedEntity.Id))
            {
                IEnumerable<IndexRelation> incomingRels = _incomingRelations[deletedEntity.Id];
                if (incomingRels != null && incomingRels.Any())
                {
                    incomingRels = incomingRels.Where(x => !x.Deleted);
                    foreach (var incomingRel in incomingRels)
                    {
                        var item = GetItem(incomingRel.From.Id);
                        if (item != null)
                        {
                            incomingRel.Deleted = true;
                            var e = new ConnectedEntityDeletedArgs(deletedEntity, incomingRel, this);
                            item.OnConnectedEntityDeleted(e);
                        }
                        if (_relations.ContainsKey(incomingRel.Id))
                        {
                            _relations.Remove(incomingRel.Id);
                        }
                    }
                }
                _incomingRelations.Remove(deletedEntity.Id);
            }
            if (_outgoingRelations.ContainsKey(deletedEntity.Id))
            {
                IEnumerable<IndexRelation> outgoingRels = _outgoingRelations[deletedEntity.Id];
                if (outgoingRels != null && outgoingRels.Any())
                {
                    outgoingRels = outgoingRels.Where(x => !x.Deleted).ToList();
                    foreach (var outgoingRel in outgoingRels)
                    {
                        var item = GetItem(outgoingRel.To.Id);
                        if (item != null)
                        {
                            outgoingRel.Deleted = true;
                            var e = new ConnectedEntityDeletedArgs(deletedEntity, outgoingRel, this);
                            item.OnConnectedEntityDeleted(e);
                        }
                        if (_relations.ContainsKey(outgoingRel.Id))
                        {
                            _relations.Remove(outgoingRel.Id);
                        }
                    }
                }
                _outgoingRelations.Remove(deletedEntity.Id);
            }
        }
    }
}
