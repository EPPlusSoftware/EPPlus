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
using OfficeOpenXml.Drawing.Vml;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Position of the a drawing.
    /// </summary>
    public class ExcelPosition : ExcelPositionBase
    {
        SetWidthCallback _setWidthCallback;
        internal ExcelPosition(XmlNamespaceManager ns, XmlNode node, SetWidthCallback setWidthCallback) :
            base(ns, node, setWidthCallback)
        {
            _setWidthCallback = setWidthCallback;
            Load();
        }
        const string colPath = "xdr:col";
        const string rowPath = "xdr:row";
        const string colOffPath = "xdr:colOff";
        const string rowOffPath = "xdr:rowOff";

        /// <summary>
        /// Load xml data
        /// </summary>
        public override void Load()
        {
            _column = GetXmlNodeInt(colPath);
            _columnOff = GetXmlNodeInt(colOffPath);
            _row = GetXmlNodeInt(rowPath);
            _rowOff = GetXmlNodeInt(rowOffPath);
        }
        /// <summary>
        /// Update xml data
        /// </summary>
        public override void UpdateXml()
        {
            SetXmlNodeString(colPath, _column.ToString());
            SetXmlNodeString(colOffPath, _columnOff.ToString());
            SetXmlNodeString(rowPath, _row.ToString());
            SetXmlNodeString(rowOffPath, _rowOff.ToString());
        }
    }
}