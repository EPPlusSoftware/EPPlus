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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// A filter column filtered by the top or botton values of an range
    /// </summary>
    public class ExcelTop10FilterColumn : ExcelFilterColumn
    {
        internal ExcelTop10FilterColumn(XmlNamespaceManager namespaceManager, XmlNode topNode) : base(namespaceManager, topNode)
        {
            FilterValue = GetXmlNodeDouble("d:top10/@filterVal");
            Percent = GetXmlNodeBool("d:top10/@percent");
            Top = GetXmlNodeBool("d:top10/@top", true);
            Value = GetXmlNodeDouble("d:top10/@val");
        }
        /// <summary>
        /// The filter value to relate to
        /// </summary>
        public double FilterValue
        {
            get;
            internal set;
        }
        /// <summary>
        /// If the filter value is an percentage
        /// </summary>
        public bool Percent
        {
            get;
            set;
        }
        /// <summary>
        /// True is top value. False is bottom values.
        /// </summary>
        public bool Top
        {
            get;
            set;
        }
        /// <summary>
        /// The value to filter on
        /// </summary>
        public double Value
        {
            get;
            set;
        }

        internal override bool Match(object value, string valueText)
        {
            if(Top)
            {
                if (Utils.ConvertUtil.IsNumericOrDate(value))
                    return Utils.ConvertUtil.GetValueDouble(value) >= FilterValue;
            }
            else
            {
                if (Utils.ConvertUtil.IsNumericOrDate(value))
                    return Utils.ConvertUtil.GetValueDouble(value) <= FilterValue;
            }
            return false;
        }

        internal override void Save()
        {
            var node = (XmlElement)CreateNode("d:top10");
            node.SetAttribute("filterVal", FilterValue.ToString("R15", CultureInfo.InvariantCulture));
            node.SetAttribute("percent", Percent ? "1": "0");
            node.SetAttribute("top", Top ? "1" : "0");
            node.SetAttribute("val", Value.ToString("R15", CultureInfo.InvariantCulture));
        }

        internal override void SetFilterValue(ExcelWorksheet worksheet, ExcelAddressBase address)
        {
            var items = new List<double>();
            var col = address._fromCol + Position;
            for (int row= address._fromRow + 1; row <= address._toRow; row++)
            {
                var v = worksheet.GetValue(row, col);
                if (Utils.ConvertUtil.IsNumericOrDate(v))
                {
                    items.Add(Utils.ConvertUtil.GetValueDouble(v));
                }
            }
            items.Sort();

            var valueInt = Convert.ToInt32(Value);
            int index;
            if (Top)
            {
                if (Percent)
                {
                    index = (items.Count - (int)((address._toRow-address._fromRow) * (valueInt / 100D)));

                }
                else
                {
                    index = items.Count - valueInt;
                }
                if (index < 0) index = 0;
                FilterValue = items[index];
            }
            else
            {
                if (Percent)
                {
                    index = (int)((address._toRow-address._fromRow) * (valueInt / 100D)) - 1;
                }
                else
                {
                    index = valueInt - 1;
                }
                if (index < 0) index = 0;
                FilterValue = index < items.Count ? items[index] : items[items.Count - 1];
            }
        }
    }
}