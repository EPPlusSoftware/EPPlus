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
