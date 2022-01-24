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
using System.Drawing;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Coloring
{
    /// <summary>
    /// Represents a system color
    /// </summary>s
    public class ExcelDrawingSystemColor : XmlHelper
    {
        internal ExcelDrawingSystemColor(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {

        }
        /// <summary>
        /// The system color
        /// </summary>
        public eSystemColor Color
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
        /// <summary>
        /// Last color computed. 
        /// </summary>
        public Color LastColor
        {
            get
            {
                return ExcelDrawingRgbColor.GetColorFromString(GetXmlNodeString("@lastClr"));
            }            
        }
        private eSystemColor TranslateFromString(string v)
        {
            switch (v)
            {
                case "btnFace":
                    return eSystemColor.ButtonFace;
                case "btnShadow":
                    return eSystemColor.ButtonShadow;
                case "btnHighlight":
                    return eSystemColor.ButtonHighlight;
                case "btnText":
                    return eSystemColor.ButtonText;
                case "3dDkShadow":
                    return eSystemColor.DarkShadow3d;
                case "3dLight":
                    return eSystemColor.Light3d;
                case "infoBk":
                    return eSystemColor.InfoBackground;
                default:
                    try
                    {
                        return (eSystemColor)Enum.Parse(typeof(eSystemColor), v, true);
                    }
                    catch
                    {
                        throw (new ArgumentException($"Invalid system color value {v}"));
                    }                    
            }
        }

        internal Color GetColor()
        {
            switch (Color)
            {
                case eSystemColor.ActiveBorder:
                    return EPPlusSystemColors.ActiveBorder;
                case eSystemColor.ActiveCaption:
                    return EPPlusSystemColors.ActiveCaption;
                case eSystemColor.AppWorkspace:
                    return EPPlusSystemColors.AppWorkspace;
                case eSystemColor.Background:
                    return EPPlusSystemColors.Window;
                case eSystemColor.ButtonFace:
                    return EPPlusSystemColors.ButtonFace;
                case eSystemColor.InactiveCaption:
                    return EPPlusSystemColors.InactiveCaption;
                case eSystemColor.Menu:
                    return EPPlusSystemColors.Menu;
                case eSystemColor.Window:
                    return EPPlusSystemColors.Window;
                case eSystemColor.WindowFrame:
                    return EPPlusSystemColors.WindowFrame;
                case eSystemColor.MenuText:
                    return EPPlusSystemColors.MenuText;
                case eSystemColor.WindowText:
                    return EPPlusSystemColors.WindowText;
                case eSystemColor.CaptionText:
                    return EPPlusSystemColors.ActiveCaptionText;
                case eSystemColor.InactiveBorder:
                    return EPPlusSystemColors.InactiveBorder;
                case eSystemColor.Highlight:
                    return EPPlusSystemColors.Highlight;
                case eSystemColor.HighlightText:
                    return EPPlusSystemColors.HighlightText;
                case eSystemColor.ButtonShadow:
                    return EPPlusSystemColors.ButtonShadow;
                case eSystemColor.GrayText:
                    return EPPlusSystemColors.GrayText;
                case eSystemColor.ButtonText:
                    return EPPlusSystemColors.ControlText;
                case eSystemColor.InactiveCaptionText:
                    return EPPlusSystemColors.InactiveCaptionText;
                case eSystemColor.ButtonHighlight:
                    return EPPlusSystemColors.ButtonHighlight;
                case eSystemColor.DarkShadow3d:
                    return EPPlusSystemColors.ControlDarkDark;
                case eSystemColor.Light3d:
                    return EPPlusSystemColors.ControlLightLight;
                case eSystemColor.InfoText:
                    return EPPlusSystemColors.InfoText;
                case eSystemColor.InfoBackground:
                    return EPPlusSystemColors.Info;
                case eSystemColor.HotLight:
                    return EPPlusSystemColors.HotTrack;
                case eSystemColor.GradientActiveCaption:
                    return EPPlusSystemColors.GradientActiveCaption;
                case eSystemColor.GradientInactiveCaption:
                    return EPPlusSystemColors.GradientInactiveCaption;
                case eSystemColor.MenuHighlight:
                    return EPPlusSystemColors.MenuHighlight;
                case eSystemColor.MenuBar:
                    return EPPlusSystemColors.MenuBar;
                default:
                    return EPPlusSystemColors.Window;
            }
        }
        private string TranslateFromEnum(eSystemColor e)
        {
            string s;
            switch (e)
            {
                case eSystemColor.ButtonFace:
                    s="btnFace";
                    break;
                case eSystemColor.ButtonShadow:
                    s = "btnShadow";
                    break;
                case eSystemColor.ButtonHighlight:
                    s = "btnHighlight";
                    break;
                case eSystemColor.ButtonText:
                    s = "btnText";
                    break;
                case eSystemColor.DarkShadow3d:
                    s = "3dDkShadow";
                    break;
                case eSystemColor.Light3d:
                    s = "3dLight";
                    break;
                case eSystemColor.InfoBackground:
                    s = "infoBk";
                    break;
                default:
                    s = e.ToString();
                    break;
            }
            return s.Substring(0, 1).ToLower()+s.Substring(1);
        }

        internal const string NodeName = "a:sysClr";
    }
}