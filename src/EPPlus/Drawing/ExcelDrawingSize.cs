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
    /// The size of the drawing 
    /// </summary>
    public class ExcelDrawingSize : XmlHelper
    {
        internal delegate void SetWidthCallback();
        SetWidthCallback _setWidthCallback;
        internal ExcelDrawingSize(XmlNamespaceManager ns, XmlNode node, SetWidthCallback setWidthCallback=null) :
            base (ns,node)
        {
            _setWidthCallback = setWidthCallback;
            Load();
        }

        private void Load()
        {
            _height = GetXmlNodeLong(colOffPath);
            _width = GetXmlNodeLong(rowOffPath);
        }
        public void UpdateXml()
        {
            SetXmlNodeString(colOffPath, _height.ToString());
            SetXmlNodeString(rowOffPath, _width.ToString());
        }
        const string colOffPath = "@cy";
        long _height=long.MinValue;
        /// <summary>
        /// Column Offset
        /// 
        /// EMU units   1cm         =   1/360000 
        ///             1US inch    =   1/914400
        ///             1pixel      =   1/9525
        /// </summary>
        public long Height
        {
            get
            {
                return _height;
            }
            set
            {
                _height = value;
                if (_setWidthCallback != null) _setWidthCallback();
            }
        }
        const string rowOffPath = "@cx";
        long _width = long.MinValue;
        /// <summary>
        /// Row Offset
        /// 
        /// EMU units   1cm         =   1/360000 
        ///             1US inch    =   1/914400
        ///             1pixel      =   1/9525
        /// </summary>
        public long Width
        {
            get
            {
                return _width;                
            }
            set
            {
                _width = value;
                if (_setWidthCallback != null) _setWidthCallback();
            }
        }
    }
}