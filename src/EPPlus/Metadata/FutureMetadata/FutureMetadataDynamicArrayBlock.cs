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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class FutureMetadataDynamicArrayBlock : FutureMetadataBlock
    {
        public FutureMetadataDynamicArrayBlock(RichDataIndexStore store, RichDataEntities entity) : base(store, RichDataEntities.FutureMetadataDynamicArrayBlock)
        {
        }

        public FutureMetadataDynamicArrayBlock(XmlReader xr, RichDataIndexStore store)
            : base(store, RichDataEntities.FutureMetadataDynamicArrayBlock)
        {
            while(!xr.EOF)
            {
                if(xr.IsElementWithName("ext"))
                {
                    Uri = xr.GetAttribute("uri");
                }
                else if(xr.IsElementWithName("dynamicArrayProperties"))
                {
                    IsDynamicArray = ConvertUtil.ToBooleanString(xr.GetAttribute("fDynamic"));
                    IsCollapsed = ConvertUtil.ToBooleanString(xr.GetAttribute("fCollapsed"));
                    ExtLstXml = xr.ReadInnerXml();
                }
                else if(xr.IsEndElementWithName("bk"))
                {
                    break;
                }
                else
                {
                    xr.Read();
                }
            }
        }

        public bool IsCollapsed { get; set; }

        public bool IsDynamicArray { get; set; }

        public string ExtLstXml { get; set; }

        public override void Save(StreamWriter sw)
        {
            sw.Write($"<bk><extLst><ext uri=\"{Uri}\"");
            if (string.IsNullOrEmpty(ExtLstXml))
            {
                sw.Write($"<xda:dynamicArrayProperties fDynamic=\"{(IsDynamicArray ? "1" : "0")}\" fCollapsed=\"{(IsCollapsed ? "1" : "0")}\"/>");
            }
            else
            {
                sw.Write($"<xda:dynamicArrayProperties fDynamic=\"{(IsDynamicArray ? "1" : "0")}\" fCollapsed=\"{(IsCollapsed ? "1" : "0")}\">");
                sw.Write($"<extLst>{ExtLstXml}</extLst>");
                sw.Write($"</xda:dynamicArrayProperties>");
            }
            sw.Write("</ext></extLst></bk>");
        }
    }
}
