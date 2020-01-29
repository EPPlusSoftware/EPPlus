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
            internal ExcelDrawingSize(XmlNamespaceManager ns, XmlNode node, SetWidthCallback setWidthCallback) :
                base (ns,node)
            {
                _setWidthCallback = setWidthCallback;
            }
            const string colOffPath = "@cy";
            /// <summary>
            /// Column Offset
            /// 
            /// EMU units   1cm         =   1/360000 
            ///             1US inch    =   1/914400
            ///             1pixel      =   1/9525
            /// </summary>
            public int Height
            {
                get
                {
                    return GetXmlNodeInt(colOffPath);
                }
                set
                {
                    SetXmlNodeString(colOffPath, value.ToString());
                    _setWidthCallback();
                }
            }
            const string rowOffPath = "@cx";
            /// <summary>
            /// Row Offset
            /// 
            /// EMU units   1cm         =   1/360000 
            ///             1US inch    =   1/914400
            ///             1pixel      =   1/9525
            /// </summary>
            public int Width
            {
                get
                {
                    return GetXmlNodeInt(rowOffPath);
                }
                set
                {
                    SetXmlNodeString(rowOffPath, value.ToString());
                    _setWidthCallback();
                }
            }
        }
}