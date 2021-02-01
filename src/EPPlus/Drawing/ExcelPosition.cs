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
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Position of the a drawing.
    /// </summary>
    public class ExcelPosition : XmlHelper
    {
        internal delegate void SetWidthCallback();
        XmlNode _node;
        XmlNamespaceManager _ns;
        SetWidthCallback _setWidthCallback;
        internal ExcelPosition(XmlNamespaceManager ns, XmlNode node, SetWidthCallback setWidthCallback) :
            base(ns, node)
        {
            _node = node;
            _ns = ns;
            _setWidthCallback = setWidthCallback;
        }
        const string colPath = "xdr:col";
        /// <summary>
        /// The column
        /// </summary>
        public int Column
        {
            get
            {
                return GetXmlNodeInt(colPath);
            }
            set
            {
                SetXmlNodeString(colPath, value.ToString());
                _setWidthCallback?.Invoke();
            }
        }
        const string rowPath = "xdr:row";
        /// <summary>
        /// The row
        /// </summary>
        public int Row
        {
            get
            {
                return GetXmlNodeInt(rowPath);
            }
            set
            {
                SetXmlNodeString(rowPath, value.ToString());
                _setWidthCallback?.Invoke();
            }
        }
        const string colOffPath = "xdr:colOff";
        /// <summary>
        /// Column Offset in EMU
        /// ss
        /// EMU units   1cm         =   1/360000 
        ///             1US inch    =   1/914400
        ///             1pixel      =   1/9525
        /// </summary>
        public int ColumnOff
        {
            get
            {
                return GetXmlNodeInt(colOffPath);
            }
            set
            {
                SetXmlNodeString(colOffPath, value.ToString());
                _setWidthCallback?.Invoke();
            }
        }
        const string rowOffPath = "xdr:rowOff";
        /// <summary>
        /// Row Offset in EMU
        /// 
        /// EMU units   1cm         =   1/360000 
        ///             1US inch    =   1/914400
        ///             1pixel      =   1/9525
        /// </summary>
        public int RowOff
        {
            get
            {
                return GetXmlNodeInt(rowOffPath);
            }
            set
            {
                SetXmlNodeString(rowOffPath, value.ToString());
                _setWidthCallback?.Invoke();
            }
        }
    }
}