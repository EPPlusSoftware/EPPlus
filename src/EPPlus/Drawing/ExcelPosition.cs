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
using System.Globalization;
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
        internal ExcelPosition(XmlNamespaceManager ns, XmlNode node, SetWidthCallback setWidthCallback, int DrawingsType = 0) :
            base(ns, node)
        {
            _setWidthCallback = setWidthCallback;
            this.excelDrawingsType = DrawingsType;
            Load();
        }

        const string colPath = "xdr:col";
        const string rowPath = "xdr:row";
        const string colOffPath = "xdr:colOff";
        const string rowOffPath = "xdr:rowOff";
        internal int _column, _row, _columnOff, _rowOff;

        int excelDrawingsType = 0;
        double _x, _y;
        const string xPath = "cdr:x";
        const string yPath = "cdr:y";

        /// <summary>
        /// X coordinate for shapes in charts.
        /// </summary>
        public double X
        {
            get
            {
                return _x;
            }
            set
            {
                _x = value;
            }
        }
        /// <summary>
        /// Y coordinate for shapes in charts.
        /// </summary>
        public double Y
        {
            get
            {
                return _y;
            }
            set
            {
                _y = value;
            }
        }

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
            if (excelDrawingsType == 0)
            {
                _column = GetXmlNodeInt(colPath);
                _columnOff = GetXmlNodeInt(colOffPath);
                _row = GetXmlNodeInt(rowPath);
                _rowOff = GetXmlNodeInt(rowOffPath);
            }
            else if (excelDrawingsType == 1)
            {
                _x = GetXmlNodeDouble(xPath);
                _y = GetXmlNodeDouble(yPath);
            }
        }
        /// <summary>
        /// Update xml data
        /// </summary>
        public void UpdateXml()
        {
            if (excelDrawingsType == 0)
            {
                SetXmlNodeString(colPath, _column.ToString());
                SetXmlNodeString(colOffPath, _columnOff.ToString());
                SetXmlNodeString(rowPath, _row.ToString());
                SetXmlNodeString(rowOffPath, _rowOff.ToString());
            }
            else if (excelDrawingsType == 1)
            {
                SetXmlNodeString(yPath, _y.ToString(CultureInfo.InvariantCulture));
                SetXmlNodeString(xPath, _x.ToString(CultureInfo.InvariantCulture));
            }
        }
    }
}