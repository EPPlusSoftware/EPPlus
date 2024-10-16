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

        public List<IndexRelation> SubRelations { get; private set; }

        public static IndexEndpoint GetSubRelationsEndpoint(RichDataIndexStore store)
        {
            var endpoint = new IndexEndpoint(store, RichDataEntities.SubRelations);
            endpoint.SubRelations = new List<IndexRelation>();
            return endpoint;
        }
    }
}
