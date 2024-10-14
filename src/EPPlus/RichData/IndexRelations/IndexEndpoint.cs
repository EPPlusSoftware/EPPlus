using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal class IndexEndpoint
    {
        public IndexEndpoint(IndexPointer pointer, int originalIndex)
        {
            _entity = pointer.EntityType;
            _id = IdGenerator.GetNewId();
            _originalIndex = originalIndex;
        }

        public IndexEndpoint(IndexedValue indexedValue, int originalIndex)
        {

        }

        private readonly RichDataEntities _entity;
        private readonly int _id;
        private readonly int _originalIndex;

        public RichDataEntities Entity => _entity;


        public int OriginalIndex => _originalIndex;

        public int CurrentIndex { get; set; }


    }
}
