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
    internal class IndexEndpoint : IdentityItem
    {
        public IndexEndpoint(RichDataIndexStore store, RichDataEntities entity)
        {
            _store = store;
            _entity = entity;
            _originalIndex = store.GetNextIndex(entity);
        }

        private IndexEndpoint(RichDataEntities entity)
        {
            _entity = entity;
        }

        private readonly RichDataEntities _entity;
        private readonly int _originalIndex;
        private readonly RichDataIndexStore _store;

        public static IndexEndpoint None = new IndexEndpoint(RichDataEntities.None);

        public RichDataEntities Entity => _entity;


        public int OriginalIndex => _originalIndex;

        public int CurrentIndex { get; set; }

        private bool _deleted;
        public bool Deleted => _deleted;

        public void DeleteMe()
        {
            _deleted = true;
            // TODO: follow relations and delete depending entities
        }

        public int? GetIndex()
        {
            return _store.GetIndexById(Id, Entity);
        }

        public virtual void InitRelations()
        {

        }

        public int? FirstTargetIndex
        {
            get
            {
                var targets = _store.GetRelationTargets(Id);
                if (targets != null && targets.Any())
                {
                    var target = targets.First();
                    return target.To.GetIndex();
                }
                return default;
            }
        }

        public int? FirstTargetId
        {
            get
            {
                var targets = _store.GetRelationTargets(Id);
                if (targets != null && targets.Any())
                {
                    var target = targets.First();
                    if (target.To.Deleted) return default;
                    return target.To.Id;
                }
                return default;
            }
        }

        public virtual T GetFirstOutgoingSubRelation<T>()
            where T : IndexEndpoint
        {
            var relations = GetOutgoingRelations(x => x.IndexType == IndexType.SubRelations);
            var entityType = _store.GetEntityByType(typeof(T));
            foreach(var relation in relations)
            {
                var parentRel = relation as IndexRelationWithSubRelations;
                if(parentRel != null && parentRel.SubRelations.Any())
                {
                    var subrel =  parentRel.SubRelations.FirstOrDefault(x => x.To.Entity == entityType);
                    if(subrel != null && subrel.To is T result)
                    {
                        return result;
                    }
                }
            }
            return default;
        }

        public IEnumerable<IndexRelation> GetOutgoingRelations()
        {
            return _store.GetRelationTargets(Id);
        }

        public IEnumerable<IndexRelation> GetOutgoingRelations(Func<IndexRelation, bool> filter)
        {
            return _store.GetRelationTargets(Id, filter);
        }

        public IEnumerable<IndexRelation> GetIncomingRelations()
        {
            return _store.GetRelationPointers(Id);
        }

        public IEnumerable<IndexRelation> GetIncomingRelations(Func<IndexRelation, bool> filter)
        {
            return _store.GetRelationPointers(Id, filter);
        }

        public T GetFirstTargetByType<T>()
            where T : IndexEndpoint
        {
            return GetFirstTargetByType<T>(x => true);
        }

        public T GetFirstTargetByType<T>(Func<IndexRelation, bool> filter)
            where T : IndexEndpoint
        {
            var entityType = _store.GetEntityByType(typeof(T));
            var targets = _store.GetRelationTargets(Id);

            if (targets != null && targets.Any())
            {
                var targetRelation = targets.FirstOrDefault(x => (filter.Invoke(x)) && x.To.Entity == entityType && !x.To.Deleted);
                if (targetRelation != null)
                {
                    return targetRelation.To as T;
                }

            }
            return default;
        }
    }
}
