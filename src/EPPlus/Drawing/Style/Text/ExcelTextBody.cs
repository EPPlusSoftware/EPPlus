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
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Properties for the textbody
    /// </summary>
    public class ExcelTextBody : XmlHelper
    {
        private readonly string _path;
        private readonly Action _initXml;
        internal ExcelTextBody(XmlNamespaceManager ns, XmlNode topNode, string path, string[] schemaNodeOrder=null, Action initXml=null) :
            base(ns, topNode)   
        {
            _path = path;
			_initXml = initXml;
			AddSchemaNodeOrder(schemaNodeOrder, new string[] { "ln", "noFill", "solidFill", "gradFill", "pattFill", "blipFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst", "highlight", "kumimoji", "lang", "altLang", "sz", "b", "i", "u", "strike", "kern", "cap", "spc", "normalizeH", "baseline", "noProof", "dirty", "err", "smtClean", "smtId", "bmk" });
        }
        /// <summary>
        /// The anchoring position within the shape
        /// </summary>
        public eTextAnchoringType Anchor
        {
            get
            {
                return GetXmlNodeString($"{_path}/@anchor").TranslateTextAchoring();
            }
            set
            {
                _initXml?.Invoke();
				SetXmlNodeString($"{_path}/@anchor", value.TranslateTextAchoringText());
            }
        }
        /// <summary>
        /// The centering of the text box.
        /// </summary>
        public bool AnchorCenter
        {
            get
            {   
                return GetXmlNodeBool($"{_path}/@anchorCtr");
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeBool($"{_path}/@anchorCtr", value, false);
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
					_initXml?.Invoke();
					SetXmlNodeString($"{_path}/@u", value.TranslateUnderlineText());
                }
            }
        }
        /// <summary>
        /// The bottom inset of the bounding rectangle
        /// </summary>
        public double? BottomInsert
        {
            get
            {
                return GetXmlNodeEmuToPtNull($"{_path}/@bIns");
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeEmuToPt($"{_path}/@bIns", value);
            }
        }
        /// <summary>
        /// The top inset of the bounding rectangle
        /// </summary>
        public double? TopInsert
        {
            get
            {
                return GetXmlNodeEmuToPtNull($"{_path}/@tIns");
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeEmuToPt($"{_path}/@tIns", value);
            }
        }
        /// <summary>
        /// The right inset of the bounding rectangle
        /// </summary>
        public double? RightInsert
        {
            get
            {
                return GetXmlNodeEmuToPtNull($"{_path}/@rIns");
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeEmuToPt($"{_path}/@rIns", value);
            }
        }
        /// <summary>
        /// The left inset of the bounding rectangle
        /// </summary>
        public double? LeftInsert
        {
            get
            {
                return GetXmlNodeEmuToPtNull($"{_path}/@lIns");
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeEmuToPt($"{_path}/@lIns", value);
            }
        }
        /// <summary>
        /// The rotation that is being applied to the text within the bounding box
        /// </summary>
        public double? Rotation
        {
            get
            {
                return GetXmlNodeAngle($"{_path}/@rot");
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeAngle($"{_path}/@rot", value, "Rotation", -100000, 100000);
            }
        }
        /// <summary>
        /// The space between text columns in the text area
        /// </summary>
        public double SpaceBetweenColumns
        {
            get
            {
                return GetXmlNodeEmuToPt($"{_path}/@spcCol");
            }
            set
            {
                if (value < 0) throw new ArgumentOutOfRangeException("SpaceBetweenColumns", "Can't be negative");
				_initXml?.Invoke();
				SetXmlNodeEmuToPt($"{_path}/@spcCol", value);
            }
        }

        /// <summary>
        /// If the before and after paragraph spacing defined by the user is to be respected
        /// </summary>
        public bool ParagraphSpacing
        {
            get
            {
                return GetXmlNodeBool($"{_path}/@spcFirstLastPara");
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeBool($"{_path}/@spcFirstLastPara", value);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public bool TextUpright
        {
            get
            {
                return GetXmlNodeBool($"{_path}/@upright");
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeBool($"{_path}/@upright", value);
            }
        }
        /// <summary>
        /// If the line spacing is decided in a simplistic manner using the font scene
        /// </summary>
        public bool CompatibleLineSpacing
        {
            get
            {
                return GetXmlNodeBool($"{_path}/@compatLnSpc");
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeBool($"{_path}/@compatLnSpc", value);
            }
        }
        /// <summary>
        /// Forces the text to be rendered anti-aliased
        /// </summary>
        public bool ForceAntiAlias
        {
            get
            {
                return GetXmlNodeBool($"{_path}/@forceAA");
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeBool($"{_path}/@forceAA", value);
            }
        }
        /// <summary>
        /// If the text within this textbox is converted from a WordArt object.
        /// </summary>
        public bool FromWordArt
        {
            get
            {
                return GetXmlNodeBool($"{_path}/@fromWordArt");
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeBool($"{_path}/@fromWordArt", value);
            }
        }
        /// <summary>
        /// If the text should be displayed vertically
        /// </summary>
        public eTextVerticalType VerticalText
        {
            get
            {
                return GetXmlNodeString($"{_path}/@vert").TranslateTextVertical();
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeString($"{_path}/@vert", value.TranslateTextVerticalText());
            }
        }
        /// <summary>
        /// If the text can flow out horizontaly
        /// </summary>
        public eTextHorizontalOverflow HorizontalTextOverflow
        {
            get
            {
                return GetXmlNodeString($"{_path}/@horzOverflow").ToEnum(eTextHorizontalOverflow.Overflow);
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeString($"{_path}/@horzOverflow", value.ToEnumString());
            }
        }

        /// <summary>
        /// If the text can flow out of the bounding box vertically
        /// </summary>
        public eTextVerticalOverflow VerticalTextOverflow
        {
            get
            {
                return GetXmlNodeString($"{_path}/@vertOverflow").ToEnum(eTextVerticalOverflow.Overflow);
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeString($"{_path}/@vertOverflow", value.ToEnumString());
            }
        }
        /// <summary>
        /// How text is wrapped
        /// </summary>
        public eTextWrappingType WrapText
        {
            get
            {
                return GetXmlNodeString($"{_path}/@wrap").ToEnum(eTextWrappingType.Square);
            }
            set
            {
				_initXml?.Invoke();
				SetXmlNodeString($"{_path}/@wrap", value.ToEnumString());
            }
        }
        /// <summary>
        /// The text within the text body should be normally auto-fited
        /// </summary>
        public eTextAutofit TextAutofit
        {
            get
            {
                if (ExistsNode($"{_path}/a:normAutofit"))
                {
                    return eTextAutofit.NormalAutofit;
                }
                else if (ExistsNode($"{_path}/a:spAutoFit"))
                {
                    return eTextAutofit.ShapeAutofit;
                }
                else
                {
                    return eTextAutofit.NoAutofit;
                }
            }
            set
            {
				_initXml?.Invoke();
				switch (value)
                {
                    case eTextAutofit.NormalAutofit:
                        if (value == TextAutofit) return;
                        DeleteNode($"{_path}/a:spAutoFit");
                        DeleteNode($"{_path}/a:noAutofit");
                        CreateNode($"{_path}/a:normAutofit");
                        break;
                    case eTextAutofit.ShapeAutofit:
                        DeleteNode($"{_path}/a:noAutofit");
                        DeleteNode($"{_path}/a:normAutofit");
                        CreateNode($"{_path}/a:spAutofit");
                        break;
                    case eTextAutofit.NoAutofit:
                        DeleteNode($"{_path}/a:spAutoFit");
                        DeleteNode($"{_path}/a:normAutofit");
                        CreateNode($"{_path}/a:noAutofit");
                        break;
                }
            }
        }
        /// <summary>
        /// The percentage of the original font size to which each run in the text body is scaled.
        /// This propery only applies when the TextAutofit property is set to NormalAutofit
        /// </summary>
        public double? AutofitNormalFontScale
        {
            get
            {
                return GetXmlNodePercentage($"{_path}/a:normAutofit/@fontScale");
            }
            set
            {
                if (TextAutofit != eTextAutofit.NormalAutofit) throw new ArgumentException("AutofitNormalFontScale", "TextAutofit must be set to NormalAutofit to use set this property");
				_initXml?.Invoke();
				SetXmlNodePercentage($"{_path}/a:normAutofit/@fontScale", value, false);
            }
        }
        /// <summary>
        /// The percentage by which the line spacing of each paragraph is reduced.
        /// This propery only applies when the TextAutofit property is set to NormalAutofit
        /// </summary>
        public double? LineSpaceReduction
        {
            get
            {
                return GetXmlNodePercentage($"{_path}/a:normAutofit/@lnSpcReduction");
            }
            set
            {
                if (TextAutofit != eTextAutofit.NormalAutofit) throw new ArgumentException("LineSpaceReduction", "TextAutofit must be set to NormalAutofit to use set this property");
				_initXml?.Invoke();
				SetXmlNodePercentage($"{_path}/a:normAutofit/@lnSpcReduction", value, false);
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
        internal void SetFromXml(XmlElement copyFromElement)
        {
            var element = PathElement;
            foreach (XmlAttribute a in copyFromElement.Attributes)
            {
                element.SetAttribute(a.Name, a.NamespaceURI, a.Value);
            }
        }
    }
}
