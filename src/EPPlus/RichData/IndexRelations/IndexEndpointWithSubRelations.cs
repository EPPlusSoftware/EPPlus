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
using OfficeOpenXml.RichData.IndexRelations.EventArguments;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal class IndexEndpointWithSubRelations : IndexEndpoint
    {
        public IndexEndpointWithSubRelations(RichDataIndexStore store, RichDataEntities entity) : base(store, entity)
        {
            _store = store;
        }

        private readonly Dictionary<RichDataEntities, IndexRelationWithSubRelations> _subRelations = new Dictionary<RichDataEntities, IndexRelationWithSubRelations>();
        private readonly RichDataIndexStore _store;

        public override IndexRelationWithSubRelations GetSubRelations(RichDataEntities entityType)
        {
            if (!_subRelations.ContainsKey(entityType)) return null;
            return _subRelations[entityType];
        }

        public override void DeleteMe(RelationDeletions relDeletions = null)
        {
            base.DeleteMe(relDeletions);
            var keys = new List<RichDataEntities>();
            foreach(var key in _subRelations.Keys)
            {
                keys.Add(key);
            }
            foreach(var key in keys)
            {
                if(_subRelations.ContainsKey(key))
                {
                    var subRel = _subRelations[key];
                    var rels = new List<IndexRelation>();
                    foreach(var rel in subRel.SubRelations)
                    {
                        rels.Add(rel);
                    }
                    foreach(var rel in rels)
                    {
                        var e = new ConnectedEntityDeletedEventArgs(rel.From, rel, _store, relDeletions);
                        rel.To.OnConnectedEntityDeleted(e);
                    }
                }
            }
            _subRelations.Clear();
        }

        public IndexRelationWithSubRelations CreateSubRelation(RichDataEntities entityType)
        {
            if(_subRelations.ContainsKey(entityType))
            {
                throw new ArgumentException($"Subrelation for entity type: {entityType} already exists.");
            }
            var subRel = _store.CreateAndAddRelationWithSubRelations(this, entityType);
            _subRelations[entityType] = subRel;
            return subRel;
        }

        protected void AddSubRelation(IndexRelation relation, RichDataEntities entityType)
        {
            if(_subRelations.ContainsKey(entityType))
            {
                var parentRel = _subRelations[entityType];
                parentRel.SubRelations.Add(relation);
            }
        }

        public override bool HasOutgoingRelationTo(RichDataEntities entityType)
        {
            if (base.HasOutgoingRelationTo(entityType)) return true;
            foreach(var et in _subRelations.Keys)
            {
                foreach(var rel in _subRelations[et].SubRelations)
                {
                    if(rel.To.EntityType == entityType) return true;
                }
            }
            return true;
        }
    }
}
