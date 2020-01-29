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
    /// Represents a special font, Complex, Latin or East asian 
    /// </summary>
    public class ExcelDrawingFontSpecial : ExcelDrawingFontBase
    {
        internal ExcelDrawingFontSpecial(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {

        }
        /// <summary>
        /// The type of font
        /// </summary>
        public eFontType Type
        {
            get
            {
                switch (TopNode.LocalName)
                {
                    case "cs":
                        return eFontType.Complex;
                    case "sym":
                        return eFontType.Symbol;
                    case "ea":
                        return eFontType.EastAsian;
                    default:
                        return eFontType.Latin;
                }
            }
            }
        /// <summary>
        /// Specifies the Panose-1 classification number for the current font using the mechanism
        /// defined in §5.2.7.17 of ISO/IEC 14496-22.
        /// This value is used as one piece of information to guide selection of a similar alternate font if the desired font is unavailable.
        /// </summary>        
        public string Panose
        {
            get
            {
                return GetXmlNodeString("@panose");
            }
            set
            {
                SetXmlNodeString("@panose",value);
            }
        }
        /// <summary>
        /// The font pitch as well as the font family for the font
        /// </summary>
        public ePitchFamily PitchFamily
        {
            get
            {
                var p=GetXmlNodeInt("@pitchFamily");
                try
                {
                    return (ePitchFamily)p;
                }
                catch
                {
                    return ePitchFamily.Default;
                }
            }
            set
            {
                SetXmlNodeString("@pitchFamily", ((int)value).ToString());
            }
        }

    }
}