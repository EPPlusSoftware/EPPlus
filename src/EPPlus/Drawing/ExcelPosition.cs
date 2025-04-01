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
    public class ExcelPosition : XmlHelper
    {
        internal delegate void SetWidthCallback();
        SetWidthCallback _setWidthCallback;
        internal ExcelPosition(XmlNamespaceManager ns, XmlNode node, SetWidthCallback setWidthCallback) :
            base(ns, node)
        {
            _setWidthCallback = setWidthCallback;
            Load();
        }
        const string colPath = "xdr:col";
        const string rowPath = "xdr:row";
        const string colOffPath = "xdr:colOff";
        const string rowOffPath = "xdr:rowOff";

        internal int _column, _row, _columnOff, _rowOff;
        /// <summary>
        /// The column
        /// </summary>
        public int Column
        {
            get
            {
                return _column;
            }
            set
            {
                _column = value;
                _setWidthCallback?.Invoke();
            }
        }
        /// <summary>
        /// The row
        /// </summary>
        public int Row
        {
            get
            {
                return _row;
            }
            set
            {
                _row = value;
                _setWidthCallback?.Invoke();
            }
        }

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
                return _columnOff;
            }
            set
            {
                _columnOff = value;
                _setWidthCallback?.Invoke();
            }
        }

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
                return _rowOff;
            }
            set
            {
                _rowOff = value;
                _setWidthCallback?.Invoke();
            }
        }

        /// <summary>
        /// Load xml data
        /// </summary>
        public void Load()
        {
            _column = GetXmlNodeInt(colPath);
            _columnOff = GetXmlNodeInt(colOffPath);
            _row = GetXmlNodeInt(rowPath);
            _rowOff = GetXmlNodeInt(rowOffPath);
        }
        /// <summary>
        /// Update xml data
        /// </summary>
        public void UpdateXml()
        {
            SetXmlNodeString(colPath, _column.ToString());
            SetXmlNodeString(colOffPath, _columnOff.ToString());
            SetXmlNodeString(rowPath, _row.ToString());
            SetXmlNodeString(rowOffPath, _rowOff.ToString());
        }
    }
}