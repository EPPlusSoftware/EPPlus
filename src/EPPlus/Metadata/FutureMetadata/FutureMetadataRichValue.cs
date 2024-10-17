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
    internal class FutureMetadataRichValue : FutureMetadataBase
    {
        public FutureMetadataRichValue(string name, RichDataIndexStore store, ExcelMetadata metadata)
            : base(store)
        {
            Name = name;
            _indexStore = store;
            Blocks = new FutureMetadataRichValueBlockCollection(store);
            var type = metadata.MetadataTypes.FirstOrDefault(t => t.Name == name);
            if(type != null)
            {
                var rel = new IndexRelation(type, this, IndexType.String);
                store.AddRelation(rel);
            }
        }
        public FutureMetadataRichValue(XmlReader xr, RichDataIndexStore store, ExcelMetadata metadata)
            : base(store)
        {
            Blocks = new FutureMetadataRichValueBlockCollection(store);
            ReadXml(xr,_indexStore, metadata);
        }

        private readonly RichDataIndexStore _indexStore;

        public override string Uri { get; set; } = ExtLstUris.RichValueDataUri;

        private void ReadXml(XmlReader xr, RichDataIndexStore store, ExcelMetadata metadata)
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
                        store.AddRelation(rel);
                    }
                    xr.Read();
                }
                else if(xr.IsElementWithName("bk"))
                {
                    Blocks.Add(new FutureMetadataRichValueBlock(xr, store));
                }
            }
        }

        public override IndexedCollection<FutureMetadataBlock> Blocks { get; set; }

        public override void Save(StreamWriter sw)
        {
            sw.Write($"<futureMetadata name=\"XLRICHVALUE\" count=\"{Blocks.Count}\">");
            foreach(var block in Blocks)
            {
                block.Save(sw);
            }
            sw.Write("</futureMetadata>");
        }
    }
}
