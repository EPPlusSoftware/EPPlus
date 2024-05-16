/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    internal class ExcelFutureMetadata
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public List<ExcelFutureMetadataType> Types { get; }=new List<ExcelFutureMetadataType>();
        //string _extLstXml;
    }
    internal abstract class ExcelFutureMetadataType
    {        
        public abstract FutureMetadataType Type { get; }
        public abstract string Uri { get; }
        public ExcelFutureMetadataDynamicArray AsDynamicArray { get { return this as ExcelFutureMetadataDynamicArray; } }
        public ExcelFutureMetadataRichData AsRichData { get { return this as ExcelFutureMetadataRichData; } }

        internal abstract void WriteXml(StreamWriter sw);
    }
    internal class ExcelFutureMetadataRichData : ExcelFutureMetadataType
    {
        public ExcelFutureMetadataRichData(int index)
        {
            Index = index;
        }
        public ExcelFutureMetadataRichData(XmlReader xr)
        {
            var startDepth = xr.Depth;
            while (xr.Read() && startDepth <= xr.Depth)
            {
                if (xr.IsElementWithName("rvb"))
                {
                    Index=int.Parse(xr.GetAttribute("i"));
                }
            }

            if (xr.NodeType == XmlNodeType.EndElement) xr.Read();
        }
        public int Index { get; private set; }
        public override FutureMetadataType Type => FutureMetadataType.RichData;
        public override string Uri => ExtLstUris.RichValueDataUri;
        internal override void WriteXml(StreamWriter sw)
        {
            sw.Write($"<xlrd:rvb i=\"{Index}\"/>");
        }
    }
    internal class ExcelFutureMetadataDynamicArray : ExcelFutureMetadataType
    {
        public ExcelFutureMetadataDynamicArray(bool isDynamicArray)
        {
            IsDynamicArray= isDynamicArray;
            IsCollapsed = false;
        }
        public ExcelFutureMetadataDynamicArray(XmlReader xr)
        {
            var startDepth = xr.Depth;
            while(xr.Read() && startDepth<=xr.Depth)
            {
                if(xr.IsElementWithName("dynamicArrayProperties"))
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
            if(string.IsNullOrEmpty(ExtLstXml))
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