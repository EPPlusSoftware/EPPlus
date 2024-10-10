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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData.Structures.SupportingPropertyBags
{
    [DebuggerDisplay("Structure: {StructureId}")]
    internal class SupportingPropertyBagData
    {
        internal string StructureId { get; set; }
        private List<SupportingPropertyBagDataItem> _items = new List<SupportingPropertyBagDataItem>();
        internal List<SupportingPropertyBagDataItem> Items => _items;

        internal static SupportingPropertyBagData CreateFromXml(XmlReader xr)
        {
            var data = new SupportingPropertyBagData();
            do
            {
                if (xr.IsElementWithName("spb"))
                {
                    data.StructureId = xr.GetAttribute("s");
                    xr.Read();
                }
                else if (xr.IsElementWithName("v"))
                {
                    var xmlSpace = xr.GetAttribute("xml:space");
                    var val = xr.ReadElementContentAsString();
                    data.Items.Add(new SupportingPropertyBagDataItem(val, xmlSpace));
                }
                else if (xr.IsEndElementWithName("spb"))
                {
                    break;
                }
                else
                {
                    xr.Read();
                }
            }
            while (!xr.EOF);
            if(data.StructureId != null && data.Items.Count > 0)
                return data;
            return null;
        }

        internal void WriteXml(StreamWriter sw)
        {
            sw.Write($"<spb s=\"{StructureId}\">");
            foreach(var item in _items)
            {
                item.WriteXml(sw);
            }
            sw.Write("</spb>");
        }
    }
}
