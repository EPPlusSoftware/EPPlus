using OfficeOpenXml.Constants;
using OfficeOpenXml.RichData;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class FutureMetadataDynamicArray : ExcelFutureMetadata
    {
        public FutureMetadataDynamicArray(bool isDynamicArray, ExcelRichData richData)
            : base(richData.IndexStore)
        {
            IsDynamicArray = isDynamicArray;
            IsCollapsed = false;
        }
        public FutureMetadataDynamicArray(XmlReader xr, ExcelRichData richData)
            : base(richData.IndexStore)
        {
            var startDepth = xr.Depth;
            while (xr.Read() && startDepth <= xr.Depth)
            {
                if (xr.IsElementWithName("dynamicArrayProperties"))
                {
                    IsDynamicArray = ConvertUtil.ToBooleanString(xr.GetAttribute("fDynamic"));
                    IsCollapsed = ConvertUtil.ToBooleanString(xr.GetAttribute("fCollapsed"));
                    ExtLstXml = xr.ReadInnerXml();
                }
            }

            if (xr.NodeType == XmlNodeType.EndElement) xr.Read();
        }

        public override string Uri { get; set; } = ExtLstUris.DynamicArrayPropertiesUri;
        public bool IsCollapsed { get; set; }
        public string ExtLstXml { get; set; }

        protected override void Save(StreamWriter sw)
        {
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
        }
    }
}
