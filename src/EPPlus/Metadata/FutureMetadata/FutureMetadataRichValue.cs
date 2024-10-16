using OfficeOpenXml.Constants;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class FutureMetadataRichValue : ExcelFutureMetadata
    {
        public FutureMetadataRichValue(string name, ExcelRichData richData, ExcelMetadata metadata)
            : base(richData.IndexStore)
        {
            Name = name;
            var type = metadata.MetadataTypes.FirstOrDefault(t => t.Name == name);
            if(type != null)
            {
                var rel = new IndexRelation(type, this, IndexType.String);
                richData.IndexStore.AddRelation(rel);
            }
        }
        public FutureMetadataRichValue(XmlReader xr, ExcelRichData richData, ExcelMetadata metadata)
            : base(richData.IndexStore)
        {
            Blocks = new FutureMetadataRichDataBlockCollection(richData);
            ReadXml(xr, richData, metadata);
        }

        public override string Uri { get; set; } = ExtLstUris.RichValueDataUri;

        private void ReadXml(XmlReader xr, ExcelRichData richData, ExcelMetadata metadata)
        {
            while(!xr.EOF)
            {
                if(xr.IsElementWithName("futureMetadata"))
                {
                    Name = xr.GetAttribute("name");
                    var type = metadata.MetadataTypes.FirstOrDefault(t => t.Name == Name);
                    if (type != null)
                    {
                        var rel = new IndexRelation(type, this, IndexType.String);
                        richData.IndexStore.AddRelation(rel);
                    }
                    xr.Read();
                }
                else if(xr.IsElementWithName("bk"))
                {
                    Blocks.Add(new FutureMetadataRichDataBlock(xr, richData));
                }
            }
        }

        public FutureMetadataRichDataBlockCollection Blocks { get; set; }
        protected override void Save(StreamWriter sw)
        {
            sw.Write("<futureMetadata name=\"XLRICHVALUE\" count=\"1\">");
            foreach(var block in Blocks)
            {
                block.Save(sw);
            }
            sw.Write("</futureMetadata>");
        }
    }
}
