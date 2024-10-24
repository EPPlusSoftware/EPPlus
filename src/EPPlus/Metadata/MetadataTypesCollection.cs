using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.IndexRelations.EventArguments;
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

        public override void Add(ExcelMetadataType item)
        {
            item.EndpointDeleted += OnMetadataTypeDeleted;
            base.Add(item);
            if(!_nameIndex.ContainsKey(item.Name))
            {
                _nameIndex.Add(item.Name, item);
            }
        }

        private void OnMetadataTypeDeleted(object source, EndpointDeletedEventArgs e)
        {
            if(source is ExcelMetadataType emt)
            {
                if(_nameIndex.ContainsKey(emt.Name))
                {
                    _nameIndex.Remove(emt.Name);
                }
            }
        }

        public override bool Remove(ExcelMetadataType item)
        {
            if(_nameIndex.ContainsKey(item.Name))
            {
                _nameIndex.Remove(item.Name);
            }
            return base.Remove(item);
        }

        public override void RemoveAt(int index)
        {
            if(index >= 0 && index < Count)
            {
                var item = this[index]; 
                if(item != null && _nameIndex.ContainsKey(item.Name)) 
                {
                    _nameIndex.Remove(item.Name);
                }
            }
            base.RemoveAt(index);
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
