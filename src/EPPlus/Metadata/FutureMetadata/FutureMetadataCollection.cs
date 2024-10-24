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
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.IndexRelations.EventArguments;
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

        public override void Add(FutureMetadataBase item)
        {
            base.Add(item);
            item.EndpointDeleted += OnEndpointDeleted;
            if(!_nameIndex.ContainsKey(item.Name))
            {
                _nameIndex[item.Name] = item;
            }
            else
            {
                _nameIndex[item.Name] = item;
            }
        }

        private void OnEndpointDeleted(object source, EndpointDeletedEventArgs e)
        {
            if(source is FutureMetadataBase fmb && _nameIndex.ContainsKey(fmb.Name))
            {
                _nameIndex.Remove(fmb.Name);
            }
        }

        public override bool Remove(FutureMetadataBase item)
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
            return false;
        }

        
    }
}
