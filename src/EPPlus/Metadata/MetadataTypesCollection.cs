using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata
{
    internal class MetadataTypesCollection : IndexedCollection<ExcelMetadataType>
    {
        public MetadataTypesCollection(RichDataIndexStore store) : base(store, RichDataEntities.MetadataType)
        {
        }

        private Dictionary<string, ExcelMetadataType> _nameIndex = new Dictionary<string, ExcelMetadataType>();

        public override RichDataEntities EntityType => RichDataEntities.MetadataType;

        public override void Add(ExcelMetadataType item)
        {
            base.Add(item);
            if(!_nameIndex.ContainsKey(item.Name))
            {
                _nameIndex.Add(item.Name, item);
            }
        }

        public bool TryGetValue(string name, out ExcelMetadataType item)
        {
            if(_nameIndex.ContainsKey(name))
            {
                item = _nameIndex[name];
                return true;
            }
            item = null;
            return false;
        }
    }
}
