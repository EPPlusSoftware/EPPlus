/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/21/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    public class ExcelControlButton : ExcelControlWithText
    {

        internal ExcelControlButton(ExcelDrawings drawings, XmlElement drawNode) : base(drawings, drawNode)
        {
        }

        internal ExcelControlButton(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackagePart part, XmlDocument controlPropertiesXml)
            : base(drawings, drawNode, control, part, controlPropertiesXml, null)
        {
        }

        public override eControlType ControlType => eControlType.Button;
        private ExcelControlMargin _margin;
        public ExcelControlMargin Margin
        {
            get
            {
                if (_margin == null)
                {
                    _margin = new ExcelControlMargin(this);
                }
                return _margin;
            }
        }        
        public eLayoutFlow LayoutFlow
        {
            get;
            set;
        }
        public eShapeOrienation Orientation
        {
            get;
            set;
        }
        public eReadingOrder ReadingOrder
        {
            get;
            set;
        }
        public bool AutomaticSize
        {
            get;
            set;
        }
        /// <summary>
        /// Text Anchoring for the text body
        /// </summary>
        internal eTextAnchoringType TextAnchor
        {
            get
            {
                return TextBody.Anchor;
            }
            set
            {
                TextBody.Anchor = value;
            }
        }
        private string _textAlignPath = "xdr:sp/xdr:txBody/a:p/a:pPr/@algn";
        /// <summary>
        /// How the text is aligned
        /// </summary>
        public eTextAlignment TextAlignment
        {
            get
            {
                switch (GetXmlNodeString(_textAlignPath))
                {
                    case "ctr":
                        return eTextAlignment.Center;
                    case "r":
                        return eTextAlignment.Right;
                    case "dist":
                        return eTextAlignment.Distributed;
                    case "just":
                        return eTextAlignment.Justified;
                    case "justLow":
                        return eTextAlignment.JustifiedLow;
                    case "thaiDist":
                        return eTextAlignment.ThaiDistributed;
                    default:
                        return eTextAlignment.Left;
                }
            }
            set
            {
                switch (value)
                {
                    case eTextAlignment.Right:
                        SetXmlNodeString(_textAlignPath, "r");
                        break;
                    case eTextAlignment.Center:
                        SetXmlNodeString(_textAlignPath, "ctr");
                        break;
                    case eTextAlignment.Distributed:
                        SetXmlNodeString(_textAlignPath, "dist");
                        break;
                    case eTextAlignment.Justified:
                        SetXmlNodeString(_textAlignPath, "just");
                        break;
                    case eTextAlignment.JustifiedLow:
                        SetXmlNodeString(_textAlignPath, "justLow");
                        break;
                    case eTextAlignment.ThaiDistributed:
                        SetXmlNodeString(_textAlignPath, "thaiDist");
                        break;
                    default:
                        DeleteNode(_textAlignPath);
                        break;
                }
            }
        }

        internal override void UpdateXml()
        {
            base.UpdateXml();
            Margin.UpdateXml();
            var vmlHelper = XmlHelperFactory.Create(_vmlProp.NameSpaceManager, _vmlProp.TopNode.ParentNode);            
            var style = "layout-flow:" + LayoutFlow.TranslateString() + ";mso-layout-flow-alt:" + Orientation.TranslateString();
            if (ReadingOrder == eReadingOrder.RightToLeft)
            {
                style += ";direction:RTL";
            }
            else if (ReadingOrder == eReadingOrder.ContextDependent)
            {
                style += ";mso-direction-alt:auto";
            }
            if(AutomaticSize)
            {
                style += ";mso-fit-shape-to-text:t";
            }
            vmlHelper.SetXmlNodeString("v:textbox/@style", style);
        }
    }
}
