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
using OfficeOpenXml.RichData.IndexRelations.EventArguments;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal class IndexEndpoint : IdentityItem
    {
        public IndexEndpoint(RichDataIndexStore store, RichDataEntities entity)
            : base(store)
        {
            _store = store;
            _entity = entity;
            //_originalIndex = store.GetNextIndex(entity);
        }

        public static readonly uint NoneId = uint.MaxValue;

        private IndexEndpoint(RichDataEntities entity)
            : base(NoneId)
        {
            if(entity != RichDataEntities.None)
            {
                throw new InvalidOperationException("This constructor can only be used for defining constant instances of IndexEndpoints");
            }
            _entity = entity;
        }

        private readonly RichDataEntities _entity;
        //private readonly int _originalIndex;
        private readonly RichDataIndexStore _store;
        
        public EventHandler<EndpointDeletedEventArgs> EndpointDeleted;

        private void OnEndpointDeleted(RelationDeletions relDeletions = null)
        {
            var args = new EndpointDeletedEventArgs(Id, relDeletions);
            EndpointDeleted?.Invoke(this, args);
        }

        public static IndexEndpoint None = new IndexEndpoint(RichDataEntities.None);

        public RichDataEntities EntityType => _entity;

        private bool _deleted;
        public bool Deleted => _deleted;

        /// <summary>
        /// Deletes an entity and its relations.
        /// </summary>
        /// <param name="relDeletions">Should not be set when called from outside the IndexRelation classes</param>
        public virtual void DeleteMe(RelationDeletions relDeletions = null)
        {
            _deleted = true;
            _store.EntityDeleted(this, relDeletions);
            OnEndpointDeleted();
        }

        public int? GetIndex()
        {
            return _store.GetIndexById(Id, EntityType);
        }

        public virtual void InitRelations()
        {

        }

        public virtual void OnConnectedEntityDeleted(ConnectedEntityDeletedEventArgs e)
        {

        }

        public int? FirstTargetIndex
        {
            get
            {
                var targets = _store.GetOutgoingRelations(Id);
                if (targets != null && targets.Any())
                {
                    var target = targets.First();
                    return target.To.GetIndex();
                }
                return default;
            }
        }

        public uint? FirstTargetId
        {
            get
            {
                var targets = _store.GetOutgoingRelations(Id);
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
            return GetFirstOutgoingSubRelation<T>(out IndexRelation r);
        }

        public virtual T GetFirstOutgoingSubRelation<T>(out IndexRelation subRelation)
            where T : IndexEndpoint
        {
            subRelation = default;
            var relations = GetOutgoingRelations(x => x.IndexType == IndexType.SubRelations);
            var entityType = _store.GetEntityByType(typeof(T));
            foreach(var relation in relations)
            {
                var parentRel = relation as IndexRelationWithSubRelations;
                if(parentRel != null && parentRel.SubRelations.Any())
                {
                    var subrel =  parentRel.SubRelations.FirstOrDefault(x => x.To.EntityType == entityType);
                    if(subrel != null && subrel.To is T result)
                    {
                        subRelation = subrel;
                        return result;
                    }
                }
            }
            return default;
        }

        public IEnumerable<IndexRelation> GetOutgoingRelations()
        {
            return _store.GetOutgoingRelations(Id);
        }

        public IEnumerable<IndexRelation> GetOutgoingRelations(Func<IndexRelation, bool> filter)
        {
            return _store.GetOutgoingRelations(Id, filter);
        }

        public IEnumerable<IndexRelation> GetIncomingRelations()
        {
            return _store.GetIncomingRelations(Id);
        }

        public IEnumerable<IndexRelation> GetIncomingRelations(Func<IndexRelation, bool> filter)
        {
            return _store.GetIncomingRelations(Id, filter);
        }

        public bool HasIncomingRelationOfType(RichDataEntities entityType)
        {
            var rels = _store.GetIncomingRelations(Id);
            foreach(var rel in rels)
            {
                if (rel.From.EntityType == entityType) return true;
            }
            return false;
        }

        public T GetFirstOutgoingRelByType<T>()
            where T : IndexEndpoint
        {
            return GetFirstOutgoingRelByType<T>(x => true);
        }

        public T GetFirstOutgoingRelByType<T>(Func<IndexRelation, bool> filter)
            where T : IndexEndpoint
        {
            var entityType = _store.GetEntityByType(typeof(T));
            var targets = _store.GetOutgoingRelations(Id);

            if (targets != null && targets.Any())
            {
                var targetRelation = targets.FirstOrDefault(x => (filter.Invoke(x)) && x.To.EntityType == entityType && !x.To.Deleted);
                if (targetRelation != null)
                {
                    return targetRelation.To as T;
                }

            }
            return default;
        }

        public T GetFirstIncomingRelByType<T>()
            where T : IndexEndpoint
        {
            return GetFirstIncomingRelByType<T>(x => true);
        }

        public T GetFirstIncomingRelByType<T>(Func<IndexRelation, bool> filter)
           where T : IndexEndpoint
        {
            var entityType = _store.GetEntityByType(typeof(T));
            var pointers = _store.GetIncomingRelations(Id);

            if (pointers != null && pointers.Any())
            {
                var pointerRelation = pointers.FirstOrDefault(x => (filter.Invoke(x)) && x.From.EntityType == entityType && !x.From.Deleted);
                if (pointerRelation != null)
                {
                    return pointerRelation.From as T;
                }

            }
            return default;
        }

        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            var otherEndpoint = obj as IndexEndpoint;
            if(otherEndpoint == null) return false;
            return otherEndpoint.Id == Id;
        }

        public override int GetHashCode()
        {
            return Id.GetHashCode();
        }

        public static bool operator ==(IndexEndpoint left, IndexEndpoint right)
        {
            if (ReferenceEquals(left, null))
            {
                return ReferenceEquals(right, null);
            }

            return left.Equals(right);
        }

        public static bool operator !=(IndexEndpoint left, IndexEndpoint right)
        {
            return !(left == right);
        }
    }
}
