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
    internal class FutureMetadataRichDataBlockCollection : IndexedCollection<FutureMetadataRichDataBlock>
    {
        public FutureMetadataRichDataBlockCollection(ExcelRichData richData) : base(richData, RichDataEntities.FutureMetadataRichDataBlock)
        {
            _richData = richData;
        }

        public FutureMetadataRichDataBlockCollection(XmlReader xr, ExcelRichData richData)
            : base(richData, RichDataEntities.FutureMetadataRichDataBlock)
        {
            _richData = richData;
        }

        private readonly ExcelRichData _richData;

        private void ReadXml(XmlReader xr)
        {
            while(!xr.EOF)
            {
                Add(new FutureMetadataRichDataBlock(xr, _richData));
                if(xr.IsEndElementWithName("futureMetadata"))
                {
                    break;
                }
                xr.Read();
            }
        }

        public override RichDataEntities EntityType => RichDataEntities.FutureMetadataRichDataBlock;

        public 
    }
}
