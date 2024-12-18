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
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.IndexRelations.EventArguments;
using OfficeOpenXml.Utils;
using System;
using System.IO;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class FutureMetadataDynamicArray : FutureMetadataBase
    {
        public FutureMetadataDynamicArray(MetadataDatabase metadataDb)
            : base(metadataDb.IndexStore)
        {
            _metadataDb = metadataDb;
            Blocks = new IndexedSubsetCollection<FutureMetadataBlock>(metadataDb.FutureMetadataRichValueBlocks);
            Blocks.CollectionIsEmpty += OnBlocksIsEmpty;
            var type = metadataDb.MetadataTypes.FirstOrDefault(t => t.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME);
            if (type != null)
            {
                type.AddRelationTo(this, IndexType.String);
            }
        }

        public FutureMetadataDynamicArray(XmlReader xr, MetadataDatabase metadataDb)
            : base(metadataDb.IndexStore)
        {
            _metadataDb = metadataDb;
            Blocks = new IndexedSubsetCollection<FutureMetadataBlock>(metadataDb.FutureMetadataRichValueBlocks);
            Blocks.CollectionIsEmpty += OnBlocksIsEmpty;
            while (!xr.EOF)
            {
                if(xr.IsElementWithName("futureMetadata"))
                {
                    Name = xr.GetAttribute("name");
                    xr.Read();
                }
                else if(xr.IsElementWithName("bk"))
                {
                    var bk = new FutureMetadataDynamicArrayBlock(xr, metadataDb.IndexStore);
                    Blocks.Add(bk);
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
            if(!string.IsNullOrEmpty(Name) && _metadataDb.MetadataTypes.TryGetValue(Name, out ExcelMetadataType type))
            {
                type.AddRelationTo(this, IndexType.String);
            }
        }

        private readonly MetadataDatabase _metadataDb;

        private void OnBlocksIsEmpty(object source, CollectionIsEmptyEventArgs e)
        {
            DeleteMe(e.Deletions);
        }

        public string ExtLstXml { get; set; }
        public override string Uri { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override IndexedSubsetCollection<FutureMetadataBlock> Blocks { get; set; }

        public static FutureMetadataDynamicArray GetDefault(MetadataDatabase metadataDb, out uint bkId)
        {
            var fm = new FutureMetadataDynamicArray(metadataDb);
            fm.Name = "XLDAPR";
            var bk = new FutureMetadataDynamicArrayBlock(metadataDb.IndexStore, RichDataEntities.FutureMetadataDynamicArrayBlock);
            bk.IsDynamicArray = true;
            bk.IsCollapsed = false;
            bkId = bk.Id;
            fm.AddRelationTo(bk);
            fm.Blocks.Add(bk);
            metadataDb.FutureMetadataDynamicArrayBlocks.Add(bk);
            return fm;
        }

        public override void Save(StreamWriter sw)
        {
            sw.Write($"<futureMetadata name=\"XLDAPR\" count=\"{Blocks.Count}\">");
            for(var x = 0; x < Blocks.Count; x++)
            {
                var block = Blocks[x];
                block.Save(sw);
            }
            sw.Write("</futureMetadata>");
        }
           
    }
}
