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
using OfficeOpenXml.Drawing.Style;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Font
{
    /// <summary>
    /// Base class a font
    /// </summary>
    public class ExcelDrawingFontBase : XmlHelper
    {
        internal ExcelDrawingFontBase(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {

        }
        /// <summary>
        /// The typeface or the name of the font
        /// </summary>
        public string Typeface
        {
            get
            {
                return GetXmlNodeString("@typeface");
            }
            internal set
            {
                SetXmlNodeString("@typeface", value);
            }
        }
    }
}