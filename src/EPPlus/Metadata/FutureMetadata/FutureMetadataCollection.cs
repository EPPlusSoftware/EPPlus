using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class FutureMetadataCollection : IndexedCollection<FutureMetadataBase>
    {
        public FutureMetadataCollection(RichDataIndexStore store) : base(store, RichDataEntities.FutureMetadata)
        {
        }

        private readonly Dictionary<string, FutureMetadataBase> _nameIndex = new Dictionary<string, FutureMetadataBase>();

        public override RichDataEntities EntityType => RichDataEntities.FutureMetadata;

        public override void Add(FutureMetadataBase item)
        {
            base.Add(item);
            if(!_nameIndex.ContainsKey(item.Name))
            {
                _nameIndex[item.Name] = item;
            }
            else
            {
                _nameIndex[item.Name] = item;
            }
        }

        public FutureMetadataBase this[string name]
        {
            get
            {
                return _nameIndex[name];
            }
            set
            {
                _nameIndex[name] = value;
            }
        }

        public bool TryGetValue(string name, out FutureMetadataBase val)
        {
            val = null;
            if(_nameIndex.ContainsKey(name))
            {
                val = _nameIndex[name];
                return true;
            }
            return true;
        }
    }
}
