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
using OfficeOpenXml.Constants;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.IndexRelations.EventArguments;
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
            Blocks = new IndexedSubsetCollection<FutureMetadataBlock>(metadata.FutureMetadataBlocks);
            Blocks.CollectionIsEmpty += OnBlocksIsEmpty;
            var type = metadata.MetadataTypes.FirstOrDefault(t => t.Name == name);
            if(type != null)
            {
                store.CreateAndAddRelation(type, this, IndexType.String);
            }
        }
        public FutureMetadataRichValue(XmlReader xr, RichDataIndexStore store, ExcelMetadata metadata)
            : base(store)
        {
            _indexStore = store;
            Blocks = new IndexedSubsetCollection<FutureMetadataBlock>(metadata.FutureMetadataBlocks);
            Blocks.CollectionIsEmpty += OnBlocksIsEmpty;
            //Blocks = new FutureMetadataRichValueBlockCollection(store);
            ReadXml(xr, metadata);
        }

        private readonly RichDataIndexStore _indexStore;


        private void OnBlocksIsEmpty(object source, CollectionIsEmptyEventArgs e)
        {
            DeleteMe(e.Deletions);
        }

        public override string Uri { get; set; } = ExtLstUris.RichValueDataUri;

        private void ReadXml(XmlReader xr, ExcelMetadata metadata)
        {
            while(!xr.EOF)
            {
                if(xr.IsElementWithName("futureMetadata"))
                {
                    Name = xr.GetAttribute("name");
                    var type = metadata.MetadataTypes.FirstOrDefault(t => t.Name == Name);
                    if (type != null)
                    {
                        _indexStore.CreateAndAddRelation( type, this, IndexType.String);
                    }
                    xr.Read();
                }
                else if(xr.IsElementWithName("bk"))
                {
                    Blocks.Add(new FutureMetadataRichValueBlock(xr, _indexStore));
                }
                else if(xr.IsEndElementWithName("futureMetadata"))
                {
                    break;
                }
                else
                {
                    xr.Read();
                }
            }
        }

        public override IndexedSubsetCollection<FutureMetadataBlock> Blocks { get; set; }

        public override void Save(StreamWriter sw)
        {
            sw.Write($"<futureMetadata name=\"XLRICHVALUE\" count=\"{Blocks.Count}\">");
            for(var ix = 0; ix < Blocks.Count; ix++)
            {
                var block = Blocks[ix];
                block.Save(sw);
            }
            sw.Write("</futureMetadata>");
        }
    }
}
