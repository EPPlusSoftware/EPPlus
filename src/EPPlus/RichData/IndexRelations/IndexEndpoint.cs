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
            _entity = entity;
            _originalIndex = store.GetNextIndex(entity);
            SubRelations = new List<IndexRelation>();
        }

        private readonly RichDataEntities _entity;
        private readonly int _originalIndex;

        public RichDataEntities Entity => _entity;


        public int OriginalIndex => _originalIndex;

        public int CurrentIndex { get; set; }

        public List<IndexRelation> SubRelations { get;}


    }
}
