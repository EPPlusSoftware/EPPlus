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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class FutureMetadataDynamicArray : FutureMetadataBase
    {
        public FutureMetadataDynamicArray(RichDataIndexStore store)
            : base(store)
        {
            Blocks = new FutureMetadataDynamicArrayBlockCollection(store);
        }
        public FutureMetadataDynamicArray(XmlReader xr, RichDataIndexStore store)
            : base(store)
        {
            Blocks = new FutureMetadataDynamicArrayBlockCollection(store);
            while (!xr.EOF)
            {
                if(xr.IsElementWithName("futureMetadata"))
                {
                    Name = xr.GetAttribute("name");
                }
                else if(xr.IsElementWithName("bk"))
                {
                    Blocks.Add(new FutureMetadataDynamicArrayBlock(xr, store));
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

            if (xr.NodeType == XmlNodeType.EndElement) xr.Read();
        }

        
        public string ExtLstXml { get; set; }
        public override string Uri { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override IndexedCollection<FutureMetadataBlock> Blocks { get; set; }

        public static FutureMetadataDynamicArray GetDefault(RichDataIndexStore store)
        {
            var fm = new FutureMetadataDynamicArray(store);
            fm.Name = "XLDAPR";
            var bk = new FutureMetadataDynamicArrayBlock(store, RichDataEntities.FutureMetadataDynamicArrayBlock);
            bk.IsDynamicArray = true;
            bk.IsCollapsed = false;
            fm.Blocks.Add(bk);
            return fm;
        }

        public override void Save(StreamWriter sw)
        {
            sw.Write($"<futureMetadata name=\"XLDAPR\" count=\"{Blocks.Count}\">");
            foreach(var block in Blocks)
            {
                block.Save(sw);
            }
            sw.Write("</futureMetadata>");
        }
           
    }
}
