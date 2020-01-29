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
using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// A filter column filtered by icons
    /// </summary>
    /// <remarks>Note that EPPlus does not filter icon columns</remarks>
    public class ExcelIconFilterColumn : ExcelFilterColumn
    {
        internal ExcelIconFilterColumn(XmlNamespaceManager namespaceManager, XmlNode topNode) : base(namespaceManager, topNode)
        {

        }
        /// <summary>
        /// The icon Id within the icon set
        /// </summary>
        public int IconId
        {
            get
            {
                return GetXmlNodeInt("d:iconId");
            }
            set
            {
                if (value < 0)
                {
                    throw (new ArgumentOutOfRangeException("iconId must not be negative"));
                }
                SetXmlNodeString("d:iconId", value.ToString(CultureInfo.InvariantCulture));
            }
        }
        /// <summary>
        /// The Iconset to filter by
        /// </summary>
        public eExcelconditionalFormattingIconsSetType IconSet
        {
            get
            {
                var v=GetXmlNodeString("d:iconSet");
                v = v.Replace("3", "Three").Replace("4", "four").Replace("5", "Five");
                try
                {
                    var r = (eExcelconditionalFormattingIconsSetType)Enum.Parse(typeof(eExcelconditionalFormattingIconsSetType), v);
                    return r;
                }
                catch
                {
                    throw (new ArgumentException($"Unknown Iconset {v}"));
                }
            }
            set
            {
                var v = value.ToString();
                v = v.Replace("Three", "3").Replace("four", "4").Replace("Five", "5");
                SetXmlNodeString("d:dxfId", v);
            }
        }

        internal override bool Match(object value, string valueText)
        {
            return true;
        }

        internal override void Save()
        {
            
        }
    }
}