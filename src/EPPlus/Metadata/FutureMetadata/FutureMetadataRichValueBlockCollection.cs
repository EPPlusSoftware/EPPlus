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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class FutureMetadataRichValueBlockCollection : IndexedCollection<FutureMetadataBlock>
    {
        public FutureMetadataRichValueBlockCollection(RichDataIndexStore store) : base(store, RichDataEntities.FutureMetadataRichDataBlock)
        {
            _store = store;
        }

        public FutureMetadataRichValueBlockCollection(XmlReader xr, RichDataIndexStore store)
            : base(store, RichDataEntities.FutureMetadataRichDataBlock)
        {
            _store = store;
        }

        private readonly RichDataIndexStore _store;

        private void ReadXml(XmlReader xr)
        {
            while(!xr.EOF)
            {
                Add(new FutureMetadataRichValueBlock(xr, _store));
                if(xr.IsEndElementWithName("futureMetadata"))
                {
                    break;
                }
                xr.Read();
            }
        }

        public override RichDataEntities EntityType => RichDataEntities.FutureMetadataRichDataBlock;

    }
}
