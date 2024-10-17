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

        public List<IndexRelation> SubRelations { get; private set; }

        public static IndexEndpoint GetSubRelationsEndpoint(RichDataIndexStore store)
        {
            var endpoint = new IndexEndpoint(store, RichDataEntities.SubRelations);
            endpoint.SubRelations = new List<IndexRelation>();
            return endpoint;
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

        public T GetFirstTargetByType<T>()
            where T : IndexEndpoint
        {
            var entityType = _store.GetEntityByType(typeof(T));
            var targets = _store.GetRelationTargets(Id);
            
            if (targets != null && targets.Any())
            {
                var target = targets.FirstOrDefault(x => x.To.Entity == entityType && !x.To.Deleted);
                if(target != null)
                {
                    return target.To as T;
                }
                
            }
            return default;
        }
    }
}
