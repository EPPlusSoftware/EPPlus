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
using System.Xml;

namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// Base class for filter columns
    /// </summary>
    public abstract class ExcelFilterColumn : XmlHelper
    {
        internal ExcelFilterColumn(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
        {

        }
        /// <summary>
        /// Gets the filter value
        /// </summary>
        /// <param name="value">The value</param>
        /// <returns></returns>
        internal protected object GetFilterValue(string value)
        {
            if ((value[0] >= '0' && value[0] <= '9') ||
                (value[value.Length - 1] >= '0' && value[value.Length - 1] <= '9'))
            {
                double d;
                if (ConvertUtil.TryParseNumericString(value, out d))
                {
                    return d;
                }
                else
                {
                    return value;
                }
            }
            else
            {
                return value;
            }
        }
        /// <summary>
        /// Zero-based index indicating the AutoFilter column to which this filter information applies
        /// </summary>
        public int Position { get => GetXmlNodeInt("@colId"); }
        const string _hiddenButtonPath = "@hiddenButton";
        /// <summary>
        /// If true the AutoFilter button for this column is hidden.
        /// </summary>
        public bool HiddenButton
        {
            get
            {
                return GetXmlNodeBool(_hiddenButtonPath);
            }
            set
            {
                SetXmlNodeBool(_hiddenButtonPath, value, false);
            }
        }
        const string _showButtonPath = "@showButton";
        /// <summary>
        /// Should filtering interface elements on this cell be shown.
        /// </summary>
        public bool ShowButton
        {
            get
            {
                return GetXmlNodeBool(_showButtonPath);
            }
            set
            {
                SetXmlNodeBool(_showButtonPath, value, true);
            }
        }

        internal abstract void Save();
        internal abstract bool Match(object value, string valueText);
        internal virtual void SetFilterValue(ExcelWorksheet worksheet, ExcelAddressBase address)
        {

        }
    }
}