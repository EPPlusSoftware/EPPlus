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
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData.Structures.SupportingPropertyBags
{
    internal class SupportingPropertyBagArray
    {
        private List<SupportingPropertyBagArrayItem> _items = new List<SupportingPropertyBagArrayItem>();

        internal List<SupportingPropertyBagArrayItem> Items => _items;

        internal static SupportingPropertyBagArray CreateFromXml(XmlReader xr)
        {
            var arr = new SupportingPropertyBagArray();
            while (!xr.EOF)
            {
                if (xr.NodeType == XmlNodeType.Element && xr.Name == "v")
                {
                    string attributeValue = xr.GetAttribute("t");
                    string elementValue = xr.ReadElementContentAsString();
                    arr.Items.Add(new SupportingPropertyBagArrayItem(attributeValue, elementValue));
                }
                else if (xr.NodeType == XmlNodeType.EndElement && xr.Name == "a")
                {
                    break;
                }
                else
                {
                    xr.Read();
                }
            }
            return arr;
        }


        internal void WriteXml(StreamWriter sw)
        {
            sw.Write($"<a count=\"{_items.Count}\">");
            foreach (var item in _items)
            {
                item.WriteXml(sw);
            }
            sw.Write("</a>");
        }
    }
}
