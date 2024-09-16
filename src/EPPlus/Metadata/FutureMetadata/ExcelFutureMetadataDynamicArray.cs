/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/25/2024         EPPlus Software AB       EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class ExcelFutureMetadataDynamicArray : ExcelFutureMetadataType
    {
        public ExcelFutureMetadataDynamicArray(bool isDynamicArray)
        {
            IsDynamicArray = isDynamicArray;
            IsCollapsed = false;
        }
        public ExcelFutureMetadataDynamicArray(XmlReader xr)
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
        internal override void WriteXml(StreamWriter sw)
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
        public override FutureMetadataType Type => FutureMetadataType.DynamicArray;
        public override string Uri => ExtLstUris.DynamicArrayPropertiesUri;
        public bool IsDynamicArray { get; set; }
        public bool IsCollapsed { get; set; }
        public string ExtLstXml { get; set; }
    }
}
