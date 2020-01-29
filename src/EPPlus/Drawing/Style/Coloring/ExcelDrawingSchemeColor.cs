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
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Coloring
{
    /// <summary>
    /// Represents a scheme color
    /// </summary>
    public class ExcelDrawingSchemeColor : XmlHelper
    {
        internal ExcelDrawingSchemeColor(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {

        }
        /// <summary>
        /// The scheme color
        /// </summary>
        public eSchemeColor Color
        {
            get
            {
                return TranslateFromString(GetXmlNodeString("@val"));
            }
            set
            {
                SetXmlNodeString("@val", TranslateFromEnum(value));
            }
        }
        private eSchemeColor TranslateFromString(string v)
        {
            switch (v.ToLower())
            {
                case "bg1":
                    return eSchemeColor.Background1;
                case "bg2":
                    return eSchemeColor.Background2;
                case "dk1":
                    return eSchemeColor.Dark1;
                case "dk2":
                    return eSchemeColor.Dark2;
                case "lt1":
                    return eSchemeColor.Light1;
                case "lt2":
                    return eSchemeColor.Light2;
                case "hlink":
                    return eSchemeColor.Hyperlink;
                case "folhlink":
                    return eSchemeColor.FollowedHyperlink;
                case "phclr":
                    return eSchemeColor.Style;
                case "tx1":
                    return eSchemeColor.Text1;
                case "tx2":
                    return eSchemeColor.Text2;
                default:
                    try
                    {
                        return (eSchemeColor)Enum.Parse(typeof(eSchemeColor), v, true);
                    }
                    catch
                    {
                        throw (new ArgumentException($"Invalid scheme color value {v}"));
                    }
            }
        }
        private string TranslateFromEnum(eSchemeColor e)
        {
            string s;
            switch (e)
            {
                case eSchemeColor.Background1:
                    s = "bg1";
                    break;
                case eSchemeColor.Background2:
                    s = "bg2";
                    break;
                case eSchemeColor.Dark1:
                    s = "dk1";
                    break;
                case eSchemeColor.Dark2:
                    s = "dk2";
                    break;
                case eSchemeColor.Light1:
                    s = "lt1";
                    break;
                case eSchemeColor.Light2:
                    s = "lt2";
                    break;
                case eSchemeColor.Hyperlink:
                    s = "hlink";
                    break;
                case eSchemeColor.FollowedHyperlink:
                    s = "folHlink";
                    break;
                case eSchemeColor.Style:
                    s = "phClr";
                    break; 
                case eSchemeColor.Text1:
                    s = "tx1";
                    break;
                case eSchemeColor.Text2:
                    s = "tx2";
                    break;
                default:
                    s = e.ToString();
                    break;
            }
            return s.Substring(0, 1).ToLower() + s.Substring(1);
        }
        internal const string NodeName = "a:schemeClr";
    }
}