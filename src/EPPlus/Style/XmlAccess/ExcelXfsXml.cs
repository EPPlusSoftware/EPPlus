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
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class xfs records. This is the top level style object.
    /// </summary>
    public sealed class ExcelXfs : StyleXmlHelper
    {
        private readonly ExcelStyles _styles;
        internal ExcelXfs(XmlNamespaceManager nameSpaceManager, ExcelStyles styles) : base(nameSpaceManager)
        {
            _styles = styles;
            isBuildIn = false;
        }
        internal ExcelXfs(XmlNamespaceManager nsm, XmlNode topNode, ExcelStyles styles) :
            base(nsm, topNode)
        {
            _styles = styles;
            XfId = GetXmlNodeInt("@xfId");
            if (XfId == 0) isBuildIn = true; //Normal taggen
            _numFmtId = GetXmlNodeInt("@numFmtId");
            FontId = GetXmlNodeInt("@fontId");
            FillId = GetXmlNodeInt("@fillId");
            BorderId = GetXmlNodeInt("@borderId");
            _readingOrder = GetReadingOrder(GetXmlNodeString(readingOrderPath));
            _indent = GetXmlNodeInt(indentPath);
            ShrinkToFit = GetXmlNodeString(shrinkToFitPath) == "1" ? true : false; 
            VerticalAlignment = GetVerticalAlign(GetXmlNodeString(verticalAlignPath));
            HorizontalAlignment = GetHorizontalAlign(GetXmlNodeString(horizontalAlignPath));
            WrapText = GetXmlNodeBool(wrapTextPath);
            _textRotation = GetXmlNodeInt(textRotationPath);
            Hidden = GetXmlNodeBool(hiddenPath);
            Locked = GetXmlNodeBool(lockedPath,true);
            QuotePrefix = GetXmlNodeBool(quotePrefixPath);
            JustifyLastLine = GetXmlNodeBool(justifyLastLine);
        }

        private ExcelReadingOrder GetReadingOrder(string value)
        {
            switch(value)
            {
                case "1":
                    return ExcelReadingOrder.LeftToRight;
                case "2":
                    return ExcelReadingOrder.RightToLeft;
                default:
                    return ExcelReadingOrder.ContextDependent;
            }
        }

        private ExcelHorizontalAlignment GetHorizontalAlign(string align)
        {
            if (align == "") return ExcelHorizontalAlignment.General;
            align = align.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + align.Substring(1, align.Length - 1);
            try
            {
                return (ExcelHorizontalAlignment)Enum.Parse(typeof(ExcelHorizontalAlignment), align);
            }
            catch
            {
                return ExcelHorizontalAlignment.General;
            }
        }

        private ExcelVerticalAlignment GetVerticalAlign(string align)
        {
            if (align == "") return ExcelVerticalAlignment.Bottom;
            align = align.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + align.Substring(1, align.Length - 1);
            try
            {
                return (ExcelVerticalAlignment)Enum.Parse(typeof(ExcelVerticalAlignment), align);
            }
            catch
            {
                return ExcelVerticalAlignment.Bottom;
            }
        }
        /// <summary>
        /// Style index
        /// </summary>
        public int XfId { get; set; }
        #region Internal Properties
        int _numFmtId;
        internal int NumberFormatId
        {
            get
            {
                return _numFmtId;
            }
            set
            {
                _numFmtId = value;
                ApplyNumberFormat = (value>0);
            }
        }

        internal int FontId { get; set; }
        internal int FillId { get; set; }
        internal int BorderId { get; set; }
        private bool isBuildIn
        {
            get;
            set;
        }
        internal bool ApplyNumberFormat
        {
            get;
            set;
        }
        internal bool ApplyFont
        {
            get;
            set;
        }
        internal bool ApplyFill
        {
            get;
            set;
        }
        internal bool ApplyBorder
        {
            get;
            set;
        }
        internal bool ApplyAlignment
        {
            get;
            set;
        }
        internal bool ApplyProtection
        {
            get;
            set;
        }
        #endregion
        #region Public Properties

        /// <summary>
        /// Numberformat properties
        /// </summary>
        public ExcelNumberFormatXml Numberformat 
        {
            get
            {
                return _styles.NumberFormats[_numFmtId < 0 ? 0 : _numFmtId];
            }
        }
        /// <summary>
        /// Font properties
        /// </summary>
        public ExcelFontXml Font 
        { 
           get
           {
               return _styles.Fonts[FontId < 0 ? 0 : FontId];
           }
        }
        /// <summary>
        /// Fill properties
        /// </summary>
        public ExcelFillXml Fill
        {
            get
            {
                return _styles.Fills[FillId < 0 ? 0 : FillId];
            }
        }        
        /// <summary>
        /// Border style properties
        /// </summary>
        public ExcelBorderXml Border
        {
            get
            {
                return _styles.Borders[BorderId < 0 ? 0 : BorderId];
            }
        }
        const string horizontalAlignPath = "d:alignment/@horizontal";

        /// <summary>
        /// Horizontal alignment
        /// </summary>
        public ExcelHorizontalAlignment HorizontalAlignment { get; set; } = ExcelHorizontalAlignment.General;
        const string verticalAlignPath = "d:alignment/@vertical";

        /// <summary>
        /// Vertical alignment
        /// </summary>
        public ExcelVerticalAlignment VerticalAlignment { get; set; } = ExcelVerticalAlignment.Bottom;
        const string justifyLastLine = "d:alignment/@justifyLastLine";
        /// <summary>
        /// If the cells justified or distributed alignment should be used on the last line of text
        /// </summary>
        public bool JustifyLastLine { get; set; } = false;
        const string wrapTextPath = "d:alignment/@wrapText";

        /// <summary>
        /// Wraped text
        /// </summary>
        public bool WrapText { get; set; } = false;
        string textRotationPath = "d:alignment/@textRotation";
        int _textRotation = 0;
        /// <summary>
        /// Text rotation angle
        /// </summary>
        public int TextRotation
        {
            get
            {
                return (_textRotation == int.MinValue ? 0 : _textRotation);
            }
            set
            {
                _textRotation = value;
            }
        }
        const string lockedPath = "d:protection/@locked";

        /// <summary>
        /// Locked when sheet is protected
        /// </summary>
        public bool Locked { get; set; } = true;
        const string hiddenPath = "d:protection/@hidden";

        /// <summary>
        /// Hide formulas when sheet is protected
        /// </summary>
        public bool Hidden { get; set; } = false;
        const string quotePrefixPath = "@quotePrefix";
        /// <summary>
        /// Prefix the formula with a quote.
        /// </summary>
        public bool QuotePrefix{ get; set; } = false;
        const string readingOrderPath = "d:alignment/@readingOrder";
        ExcelReadingOrder _readingOrder = ExcelReadingOrder.ContextDependent;
        /// <summary>
        /// Readingorder
        /// </summary>
        public ExcelReadingOrder ReadingOrder
        {
            get
            {
                return _readingOrder;
            }
            set
            {
                _readingOrder = value;
            }
        }
        const string shrinkToFitPath = "d:alignment/@shrinkToFit";

        /// <summary>
        /// Shrink to fit
        /// </summary>
        public bool ShrinkToFit { get; set; } = false;
        const string indentPath = "d:alignment/@indent";
        int _indent = 0;
        /// <summary>
        /// Indentation
        /// </summary>
        public int Indent
        {
            get
            {
                return (_indent == int.MinValue ? 0 : _indent);
            }
            set
            {
                _indent=value;
            }
        }
        #endregion
        internal void RegisterEvent(ExcelXfs xf)
        {
            //                RegisterEvent(xf, xf.Xf_ChangedEvent);
        }
        internal override string Id
        {

            get
            {
                return XfId + "|" + NumberFormatId.ToString() + "|" + FontId.ToString() + "|" + FillId.ToString() + "|" + BorderId.ToString() + VerticalAlignment.ToString() + "|" + HorizontalAlignment.ToString() + "|" + WrapText.ToString() + "|" + ReadingOrder.ToString() + "|" + isBuildIn.ToString() + TextRotation.ToString() + Locked.ToString() + Hidden.ToString() + ShrinkToFit.ToString() + Indent.ToString() + QuotePrefix.ToString() + JustifyLastLine.ToString(); 
                //return Numberformat.Id + "|" + Font.Id + "|" + Fill.Id + "|" + Border.Id + VerticalAlignment.ToString() + "|" + HorizontalAlignment.ToString() + "|" + WrapText.ToString() + "|" + ReadingOrder.ToString(); 
            }
        }
        internal ExcelXfs Copy()
        {
            return Copy(_styles);
        }        
        internal ExcelXfs Copy(ExcelStyles styles)
        {
            ExcelXfs newXF = new ExcelXfs(NameSpaceManager, styles);
            newXF.NumberFormatId = _numFmtId;
            newXF.FontId = FontId;
            newXF.FillId = FillId;
            newXF.BorderId = BorderId;
            newXF.XfId = XfId;
            newXF.ReadingOrder = _readingOrder;
            newXF.HorizontalAlignment = HorizontalAlignment;
            newXF.VerticalAlignment = VerticalAlignment;
            newXF.WrapText = WrapText;
            newXF.ShrinkToFit = ShrinkToFit;
            newXF.Indent = _indent;
            newXF.TextRotation = _textRotation;
            newXF.Locked = Locked;
            newXF.Hidden = Hidden;
            newXF.QuotePrefix = QuotePrefix;
            newXF.JustifyLastLine = JustifyLastLine;
            return newXF;
        }

        internal int GetNewID(ExcelStyleCollection<ExcelXfs> xfsCol, StyleBase styleObject, eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            ExcelXfs newXfs = this.Copy();
            switch(styleClass)
            {
                case eStyleClass.Numberformat:
                    newXfs.NumberFormatId = GetIdNumberFormat(styleProperty, value);
                    styleObject.SetIndex(newXfs.NumberFormatId);
                    break;
                case eStyleClass.Font:
                {
                    newXfs.FontId = GetIdFont(styleProperty, value);
                    styleObject.SetIndex(newXfs.FontId);
                    break;
                }
                case eStyleClass.Fill:
                case eStyleClass.FillBackgroundColor:
                case eStyleClass.FillPatternColor:
                    newXfs.FillId = GetIdFill(styleClass, styleProperty, value);
                    styleObject.SetIndex(newXfs.FillId);
                    break;
                case eStyleClass.GradientFill:
                case eStyleClass.FillGradientColor1:
                case eStyleClass.FillGradientColor2:
                    newXfs.FillId = GetIdGradientFill(styleClass, styleProperty, value);
                    styleObject.SetIndex(newXfs.FillId);
                    break;
                case eStyleClass.Border:
                case eStyleClass.BorderBottom:
                case eStyleClass.BorderDiagonal:
                case eStyleClass.BorderLeft:
                case eStyleClass.BorderRight:
                case eStyleClass.BorderTop:
                    newXfs.BorderId = GetIdBorder(styleClass, styleProperty, value);
                    styleObject.SetIndex(newXfs.BorderId);
                    break;
                case eStyleClass.Style:
                    switch(styleProperty)
                    {
                        case eStyleProperty.XfId:
                            newXfs.XfId = (int)value;
                            break;
                        case eStyleProperty.HorizontalAlign:
                            newXfs.HorizontalAlignment=(ExcelHorizontalAlignment)value;
                            break;
                        case eStyleProperty.VerticalAlign:
                            newXfs.VerticalAlignment = (ExcelVerticalAlignment)value;
                            break;
                        case eStyleProperty.WrapText:
                            newXfs.WrapText = (bool)value;
                            break;
                        case eStyleProperty.ReadingOrder:
                            newXfs.ReadingOrder = (ExcelReadingOrder)value;
                            break;
                        case eStyleProperty.ShrinkToFit:
                            newXfs.ShrinkToFit=(bool)value;
                            break;
                        case eStyleProperty.Indent:
                            newXfs.Indent = (int)value;
                            break;
                        case eStyleProperty.TextRotation:
                            newXfs.TextRotation = (int)value;
                            break;
                        case eStyleProperty.Locked:
                            newXfs.Locked = (bool)value;
                            break;
                        case eStyleProperty.Hidden:
                            newXfs.Hidden = (bool)value;
                            break;
                        case eStyleProperty.QuotePrefix:
                            newXfs.QuotePrefix = (bool)value;
                            break;
                        case eStyleProperty.JustifyLastLine:
                            newXfs.JustifyLastLine = (bool)value;
                            break;
                        default:
                            throw (new Exception("Invalid property for class style."));

                    }
                    break;
                default:
                    break;
            }
            int id = xfsCol.FindIndexById(newXfs.Id);
            if (id < 0)
            {
                return xfsCol.Add(newXfs.Id, newXfs);
            }
            return id;
        }

        private int GetIdBorder(eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            ExcelBorderXml border = Border.Copy();

            switch (styleClass)
            {
                case eStyleClass.BorderBottom:
                    SetBorderItem(border.Bottom, styleProperty, value);
                    break;
                case eStyleClass.BorderDiagonal:
                    SetBorderItem(border.Diagonal, styleProperty, value);
                    break;
                case eStyleClass.BorderLeft:
                    SetBorderItem(border.Left, styleProperty, value);
                    break;
                case eStyleClass.BorderRight:
                    SetBorderItem(border.Right, styleProperty, value);
                    break;
                case eStyleClass.BorderTop:
                    SetBorderItem(border.Top, styleProperty, value);
                    break;
                case eStyleClass.Border:
                    if (styleProperty == eStyleProperty.BorderDiagonalUp)
                    {
                        border.DiagonalUp = (bool)value;
                    }
                    else if (styleProperty == eStyleProperty.BorderDiagonalDown)
                    {
                        border.DiagonalDown = (bool)value;
                    }
                    else
                    {
                        throw (new Exception("Invalid property for class Border."));
                    }
                    break;
                default:
                    throw (new Exception("Invalid class/property for class Border."));
            }
            int subId;
            string id = border.Id;
            subId = _styles.Borders.FindIndexById(id);
            if (subId == int.MinValue)
            {
                return _styles.Borders.Add(id, border);
            }
            return subId;
        }

        private void SetBorderItem(ExcelBorderItemXml excelBorderItem, eStyleProperty styleProperty, object value)
        {
            if(styleProperty==eStyleProperty.Style)
            {
                excelBorderItem.Style = (ExcelBorderStyle)value;
                return;
            }

            //Check that we have an style
            if (excelBorderItem.Style == ExcelBorderStyle.None)
            {
                throw (new InvalidOperationException("Can't set bordercolor when style is not set."));
            }

            if (styleProperty == eStyleProperty.Color)
            {
                excelBorderItem.Color.Rgb = value.ToString();
            }
            else if(styleProperty == eStyleProperty.Theme)
            {
                excelBorderItem.Color.Theme = (eThemeSchemeColor?)value;
            }
            else if (styleProperty == eStyleProperty.IndexedColor)
            {
                excelBorderItem.Color.Indexed = (int)value;
            }
            else if (styleProperty == eStyleProperty.Tint)
            {
                excelBorderItem.Color.Tint = (decimal)value;
            }
            else if (styleProperty == eStyleProperty.AutoColor)
            {
                excelBorderItem.Color.Auto = (bool)value;
            }
        }

        private int GetIdFill(eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            ExcelFillXml fill = Fill.Copy();

            switch (styleProperty)
            {
                case eStyleProperty.PatternType:
                    if (fill is ExcelGradientFillXml)
                    {
                        fill = new ExcelFillXml(NameSpaceManager);
                    }
                    fill.PatternType = (ExcelFillStyle)value;
                    break;
                case eStyleProperty.Color:
                case eStyleProperty.Tint:
                case eStyleProperty.IndexedColor:
                case eStyleProperty.AutoColor:
                case eStyleProperty.Theme:
                    if (fill is ExcelGradientFillXml)
                    {
                        fill = new ExcelFillXml(NameSpaceManager);
                    }
                    if (fill.PatternType == ExcelFillStyle.None)
                    {
                        throw (new ArgumentException("Can't set color when patterntype is not set."));
                    }
                    ExcelColorXml destColor;
                    if (styleClass==eStyleClass.FillPatternColor)
                    {
                        destColor = fill.PatternColor;
                    }
                    else
                    {
                        destColor = fill.BackgroundColor;
                    }

                    if (styleProperty == eStyleProperty.Color)
                    {
                        destColor.Rgb = value.ToString();
                    }
                    else if (styleProperty == eStyleProperty.Tint)
                    {
                        destColor.Tint = (decimal)value;
                    }
                    else if (styleProperty == eStyleProperty.IndexedColor)
                    {
                        destColor.Indexed = (int)value;
                    }
                    else if(styleProperty == eStyleProperty.Theme)
                    {
                        destColor.Theme = (eThemeSchemeColor?)value;
                    }
                    else
                    {
                        destColor.Auto = (bool)value;
                    }

                    break;
                default:
                    throw (new ArgumentException("Invalid class/property for class Fill."));
            }
            int subId;
            string id = fill.Id;
            subId = _styles.Fills.FindIndexById(id);
            if (subId == int.MinValue)
            {
                return _styles.Fills.Add(id, fill);
            }
            return subId;
        }
        private int GetIdGradientFill(eStyleClass styleClass, eStyleProperty styleProperty, object value)
        {
            ExcelGradientFillXml fill;
            if(Fill is ExcelGradientFillXml)
            {
                fill = (ExcelGradientFillXml)Fill.Copy();
            }
            else
            {
                fill = new ExcelGradientFillXml(Fill.NameSpaceManager);
                fill.GradientColor1.SetColor(Color.White);
                fill.GradientColor2.SetColor(Color.FromArgb(79,129,189));
                fill.Type=ExcelFillGradientType.Linear;
                fill.Degree=90;
                fill.Top = double.NaN;
                fill.Bottom = double.NaN;
                fill.Left = double.NaN;
                fill.Right = double.NaN;
            }

            switch (styleProperty)
            {
                case eStyleProperty.GradientType:
                    fill.Type = (ExcelFillGradientType)value;
                    break;
                case eStyleProperty.GradientDegree:
                    fill.Degree = (double)value;
                    break;
                case eStyleProperty.GradientTop:
                    fill.Top = (double)value;
                    break;
                case eStyleProperty.GradientBottom: 
                    fill.Bottom = (double)value;
                    break;
                case eStyleProperty.GradientLeft:
                    fill.Left = (double)value;
                    break;
                case eStyleProperty.GradientRight:
                    fill.Right = (double)value;
                    break;
                case eStyleProperty.Color:
                case eStyleProperty.Tint:
                case eStyleProperty.IndexedColor:
                case eStyleProperty.AutoColor:
                case eStyleProperty.Theme:
                    ExcelColorXml destColor;

                    if (styleClass == eStyleClass.FillGradientColor1)
                    {
                        destColor = fill.GradientColor1;
                    }
                    else
                    {
                        destColor = fill.GradientColor2;
                    }
                    
                    if (styleProperty == eStyleProperty.Color)
                    {
                        destColor.Rgb = value.ToString();
                    }
                    else if (styleProperty == eStyleProperty.Tint)
                    {
                        destColor.Tint = (decimal)value;
                    }
                    else if (styleProperty == eStyleProperty.Theme)
                    {
                        destColor.Theme = (eThemeSchemeColor?)value;
                    }
                    else if (styleProperty == eStyleProperty.IndexedColor)
                    {
                        destColor.Indexed = (int)value;
                    }
                    else
                    {
                        destColor.Auto = (bool)value;
                    }
                    break;
                default:
                    throw (new ArgumentException("Invalid class/property for class Fill."));
            }
            int subId;
            string id = fill.Id;
            subId = _styles.Fills.FindIndexById(id);
            if (subId == int.MinValue)
            {
                return _styles.Fills.Add(id, fill);
            }
            return subId;
        }

        private int GetIdNumberFormat(eStyleProperty styleProperty, object value)
        {
            if (styleProperty == eStyleProperty.Format)
            {
                ExcelNumberFormatXml item=null;
                if (!_styles.NumberFormats.FindById(value.ToString(), ref item))
                {
                    item = new ExcelNumberFormatXml(NameSpaceManager) { Format = value.ToString(), NumFmtId = _styles.NumberFormats.NextId++ };
                    _styles.NumberFormats.Add(value.ToString(), item);
                }
                return item.NumFmtId;
            }
            else
            {
                throw (new Exception("Invalid property for class Numberformat"));
            }
        }
        private int GetIdFont(eStyleProperty styleProperty, object value)
        {
            ExcelFontXml fnt = Font.Copy();

            switch (styleProperty)
            {
                case eStyleProperty.Name:
                    fnt.Name = value.ToString();
                    break;
                case eStyleProperty.Size:
                    fnt.Size = (float)value;
                    break;
                case eStyleProperty.Family:
                    fnt.Family = (int)value;
                    break;
                case eStyleProperty.Bold:
                    fnt.Bold = (bool)value;
                    break;
                case eStyleProperty.Italic:
                    fnt.Italic = (bool)value;
                    break;
                case eStyleProperty.Strike:
                    fnt.Strike = (bool)value;
                    break;
                case eStyleProperty.UnderlineType:
                    fnt.UnderLineType = (ExcelUnderLineType)value;
                    break;
                case eStyleProperty.Color:
                    fnt.Color.Rgb=value.ToString();
                    break;
                case eStyleProperty.Tint:
                    fnt.Color.Tint = (decimal)value;
                    break;
                case eStyleProperty.Theme:
                    fnt.Color.Theme = (eThemeSchemeColor?)value;
                    break;
                case eStyleProperty.IndexedColor:
                    fnt.Color.Indexed = (int)value;
                    break;
                case eStyleProperty.AutoColor:
                    fnt.Color.Auto = (bool)value;
                    break;
                case eStyleProperty.VerticalAlign:
                    fnt.VerticalAlign = ((ExcelVerticalAlignmentFont)value) == ExcelVerticalAlignmentFont.None ? "" : value.ToString().ToLower(CultureInfo.InvariantCulture);
                    break;
                case eStyleProperty.Scheme:
                    fnt.Scheme = value.ToString();
                    break;
                case eStyleProperty.Charset:
                    fnt.Charset = (int?)value;
                    break;
                default:
                    throw (new Exception("Invalid property for class Font"));
            }
            int subId;
            string id = fnt.Id;
            subId = _styles.Fonts.FindIndexById(id);
            if (subId == int.MinValue)
            {
                return _styles.Fonts.Add(id,fnt);
            }
            return subId;
        }
        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            return CreateXmlNode(topNode, false);
        }
        internal XmlNode CreateXmlNode(XmlNode topNode, bool isCellStyleXsf)
        {
            TopNode = topNode;
            if(XfId<0 || XfId>=_styles.CellStyleXfs.Count) //XfId has an invalid reference. Remove it.
            {
                XfId = int.MinValue;
            }

            var doSetXfId = (!isCellStyleXsf && XfId > int.MinValue && _styles.CellStyleXfs.Count > 0 && _styles.CellStyleXfs[XfId].newID >= 0);
            if (_numFmtId >= 0)
            {
                SetXmlNodeString("@numFmtId", _numFmtId.ToString());
                if(_numFmtId > 0) SetXmlNodeString("@applyNumberFormat", "1");
            }
            if (FontId >= 0)
            {
                SetXmlNodeString("@fontId", _styles.Fonts[FontId].newID.ToString());
                if(FontId > 0) SetXmlNodeString("@applyFont", "1");
            }
            if (FillId >= 0)
            {
                SetXmlNodeString("@fillId", _styles.Fills[FillId].newID.ToString());
                if(FillId > 0) SetXmlNodeString("@applyFill", "1");
            }
            if (BorderId >= 0)
            {
                SetXmlNodeString("@borderId", _styles.Borders[BorderId].newID.ToString());
                if(BorderId > 0) SetXmlNodeString("@applyBorder", "1");
            }
            if(HorizontalAlignment != ExcelHorizontalAlignment.General) SetXmlNodeString(horizontalAlignPath, SetAlignString(HorizontalAlignment));
            if (doSetXfId)
            {
                SetXmlNodeString("@xfId", _styles.CellStyleXfs[XfId].newID.ToString());
            }

            if(VerticalAlignment != ExcelVerticalAlignment.Bottom) SetXmlNodeString(verticalAlignPath, SetAlignString(VerticalAlignment));
            if(WrapText) SetXmlNodeString(wrapTextPath, "1");
            if(_readingOrder!=ExcelReadingOrder.ContextDependent) SetXmlNodeString(readingOrderPath, ((int)_readingOrder).ToString());
            if(ShrinkToFit) SetXmlNodeString(shrinkToFitPath, "1");
            if(_indent > 0) SetXmlNodeString(indentPath, _indent.ToString());
            if(_textRotation > 0) SetXmlNodeString(textRotationPath, _textRotation.ToString());
            if(!Locked) SetXmlNodeString(lockedPath, "0");
            if(Hidden) SetXmlNodeString(hiddenPath, "1");
            if(QuotePrefix) SetXmlNodeString(quotePrefixPath, "1");
            if(JustifyLastLine) SetXmlNodeString(justifyLastLine, "1");

            if ((Locked == false || Hidden == false))
            {
                SetXmlNodeString("@applyProtection", "1");
            }

            if ((HorizontalAlignment != ExcelHorizontalAlignment.General || VerticalAlignment != ExcelVerticalAlignment.Bottom))
            {
                SetXmlNodeString("@applyAlignment", "1");
            }

            return TopNode;
        }

        private string SetAlignString(Enum align)
        {
            string newName = Enum.GetName(align.GetType(), align);
            return newName.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + newName.Substring(1, newName.Length - 1);
        }
    }
}