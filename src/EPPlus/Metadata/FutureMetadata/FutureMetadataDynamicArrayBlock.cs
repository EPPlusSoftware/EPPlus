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
