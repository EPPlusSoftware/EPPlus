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
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// Represents a value filter column
    /// </summary>
    public class ExcelValueFilterColumn : ExcelFilterColumn
    {
        internal ExcelValueFilterColumn(XmlNamespaceManager namespaceManager, XmlNode topNode) : base(namespaceManager, topNode)
        {
            Filters = new ExcelValueFilterCollection();
            LoadFilters(topNode);
        }

        private void LoadFilters(XmlNode topNode)
        {
            foreach (XmlNode node in topNode.FirstChild.ChildNodes)
            {
                switch (node.LocalName)
                {
                    case "filter":
                        Filters.Add(new ExcelFilterValueItem(node.Attributes["val"].Value));
                        break;
                    case "dateGroupItem":
                        var item = CreateDateGroupItem(node);
                        if (item != null)
                        {
                            Filters.Add(item);
                        }
                        break;
                }
            }
        }

        private ExcelFilterDateGroupItem CreateDateGroupItem(XmlNode node)
        {
            try
            {
                var xml=XmlHelperFactory.Create(NameSpaceManager, node);
                var grouping = (eDateTimeGrouping)Enum.Parse(typeof(eDateTimeGrouping), xml.GetXmlNodeString("@dateTimeGrouping"), true);
                switch (grouping)
                {
                    case eDateTimeGrouping.Year:
                        return new ExcelFilterDateGroupItem(xml.GetXmlNodeInt("@year"));
                    case eDateTimeGrouping.Month:
                        return new ExcelFilterDateGroupItem(xml.GetXmlNodeInt("@year"), xml.GetXmlNodeInt("@month"));
                    case eDateTimeGrouping.Day:
                        return new ExcelFilterDateGroupItem(xml.GetXmlNodeInt("@year"), xml.GetXmlNodeInt("@month"), xml.GetXmlNodeInt("@day"));
                    case eDateTimeGrouping.Hour:
                        return new ExcelFilterDateGroupItem(xml.GetXmlNodeInt("@year"), xml.GetXmlNodeInt("@month"), xml.GetXmlNodeInt("@day"), xml.GetXmlNodeInt("@hour"));
                    case eDateTimeGrouping.Minute:
                        return new ExcelFilterDateGroupItem(xml.GetXmlNodeInt("@year"), xml.GetXmlNodeInt("@month"), xml.GetXmlNodeInt("@day"), xml.GetXmlNodeInt("@hour"), xml.GetXmlNodeInt("@minute"));
                    default:
                        return new ExcelFilterDateGroupItem(xml.GetXmlNodeInt("@year"), xml.GetXmlNodeInt("@month"), xml.GetXmlNodeInt("@day"), xml.GetXmlNodeInt("@hour"), xml.GetXmlNodeInt("@minute"), xml.GetXmlNodeInt("@second"));
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// The filters applied to the columns
        /// </summary>
        public ExcelValueFilterCollection Filters { get; set; }
        internal override bool Match(object value, string valueText)
        {
            var match = true;
            foreach (var filter in Filters)
            {
                if(filter is ExcelFilterDateGroupItem d)
                {
                    var valueDate = Utils.ConvertUtil.GetValueDate(value);
                    match = valueDate.HasValue && d.Match(valueDate.Value);                    
                }
                else if (filter is ExcelFilterValueItem v)
                {
                    if(string.IsNullOrEmpty(valueText))
                    {
                        match = Filters.Blank;
                    }
                    else
                    {
                        match = v.Value == valueText;
                    }
                }
                if (match) return true;
            }
            return match;
        }
        internal override void Save()
        {
            var node = (XmlElement)CreateNode("d:filters");
            node.RemoveAll();
            if (Filters.Blank) (node).SetAttribute("blank", "1");
            if (Filters.CalendarType.HasValue) (node).SetAttribute("calendarType", Filters.CalendarType.Value.ToEnumString());
            foreach(var f in Filters)
            {
                if(f is ExcelFilterDateGroupItem d)
                {
                    d.AddNode(node);
                }
                else
                {
                    var e = TopNode.OwnerDocument.CreateElement("filter", ExcelPackage.schemaMain);
                    e.SetAttribute("val", ((ExcelFilterValueItem)f).Value);
                    node.AppendChild(e);
                }
            }
        }

        private string ConvertToString(object f)
        {
            return f?.ToString();
        }
    }
}