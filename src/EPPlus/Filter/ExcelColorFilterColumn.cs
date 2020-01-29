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
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// Represents a column filtered by colors.
    /// 
    /// </summary>
    public class ExcelColorFilterColumn : ExcelFilterColumn
    {
        internal ExcelColorFilterColumn(XmlNamespaceManager namespaceManager, XmlNode topNode) : base(namespaceManager, topNode)
        {

        }
        /// <summary>
        /// Indicating whether or not to filter by the cell's fill color. 
        /// True filters by cell fill. 
        /// False filter by the cell's font color.
        /// </summary>
        public bool CellColor
        {
            get
            {
                return GetXmlNodeBool("d:cellColor");
            }
            set
            {
                SetXmlNodeBool("d:cellColor", value);
            }
        }
        /// <summary>
        /// The differencial Style Id, referencing the DXF styles collection
        /// </summary>
        public int DxfId
        {
            get
            {
                return GetXmlNodeInt("d:dxfId");
            }
            set
            {
                if(value<0)
                {
                    throw (new ArgumentOutOfRangeException("DfxId must not be negative"));
                }
                SetXmlNodeString("d:dxfId", value.ToString(CultureInfo.InvariantCulture));
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