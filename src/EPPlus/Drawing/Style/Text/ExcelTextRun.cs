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
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.Extensions;
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// A richtext part
    /// </summary>
    public class ExcelTextRun : XmlHelper
    {
        string _path;
        internal ExcelTextRun(XmlNamespaceManager ns, XmlNode topNode, string path) :
            base(ns, topNode)
        {
            _path = path;
            SchemaNodeOrder = new string[] { "ln", "noFill", "solidFill", "gradFill", "pattFill", "blipFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst", "highlight", "kumimoji", "lang", "altLang", "sz", "b", "i", "u", "strike", "kern", "cap", "spc", "normalizeH", "baseline", "noProof", "dirty", "err", "smtClean", "smtId", "bmk" };
        }
        /// <summary>
        /// The capitalization that is to be applied
        /// </summary>
        public eTextCapsType Capitalization
        {
            get
            {
                return GetXmlNodeString($"{_path}/@cap").ToEnum(eTextCapsType.None);
            }
            set
            {
                SetXmlNodeString($"{_path}/@kern", value.ToEnumString());
            }
        }
        /// <summary>
        /// The minimum font size at which character kerning occurs
        /// </summary>
        public double Kerning
        {
            get
            {
                return GetXmlNodeFontSize($"{_path}/@kern");
            }
            set
            {
                SetXmlNodeFontSize($"{_path}/@kern", value, "Kerning");
            }
        }
        /// <summary>
        /// Fontsize
        /// Spans from 0-4000
        /// </summary>
        public double FontSize
        {
            get
            {
                return GetXmlNodeFontSize($"{_path}/@sz");
            }
            set
            {
                SetXmlNodeFontSize($"{_path}/@sz", value, "FontSize");
            }
        }
        /// <summary>
        /// The spacing between between characters
        /// </summary>
        public double Spacing
        {
            get
            {
                return GetXmlNodeFontSize($"{_path}/@spc");
            }
            set
            {
                SetXmlNodeFontSize($"{_path}/@spc", value, "Spacing", true);
            }
        }
        /// <summary>
        /// The baseline for both the superscript and subscript fonts in percentage
        /// </summary>
        public double Baseline
        {
            get
            {
                return GetXmlNodePercentage($"{_path}/@baseline")??0;
            }
            set
            {
                SetXmlNodePercentage($"{_path}/@baseline", value);
            }
        }
        /// <summary>
        /// Bold text
        /// </summary>
        public bool Bold
        {
            get
            {
                return GetXmlNodeBool($"{_path}/@b");
            }
            set
            {
                SetXmlNodeBool($"{_path}/@b", value, false);
            }
        }
        /// <summary>
        /// Italic text
        /// </summary>
        public bool Italic
        {
            get
            {
                return GetXmlNodeBool($"{_path}/@i");
            }
            set
            {
                SetXmlNodeBool($"{_path}/@i", value, false);
            }
        }
        /// <summary>
        /// Strike-out text
        /// </summary>
        public eStrikeType Strike
        {
            get
            {
                return GetXmlNodeString($"{_path}/@strike").TranslateStrikeType();
            }
            set
            {
                SetXmlNodeString($"{_path}/@strike", value.TranslateStrikeTypeText());
            }
        }
        /// <summary>
        /// Underlined text
        /// </summary>
        public eUnderLineType UnderLine
        {
            get
            {
                return GetXmlNodeString($"{_path}/@u").TranslateUnderline();
            }
            set
            {
                if (value == eUnderLineType.None) 
                {
                    DeleteNode($"{_path}/@u");
                }
                else
                {
                    SetXmlNodeString($"{_path}/@u", value.TranslateUnderlineText());
                }
            }
        }
        internal XmlElement PathElement
        {
            get
            {
                var node = (XmlElement)GetNode(_path);
                if (node == null)
                {
                    return (XmlElement)CreateNode(_path);
                }
                else
                {
                    return node;
                }
            }
        }
   }
}
