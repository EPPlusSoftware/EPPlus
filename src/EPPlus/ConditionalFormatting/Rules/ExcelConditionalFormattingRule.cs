using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Xml;
using OfficeOpenXml.Utils.Extensions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Linq;
using OfficeOpenXml.Style;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System.Drawing;
using OfficeOpenXml.Packaging;
using System.Globalization;

namespace OfficeOpenXml.ConditionalFormatting
{
    public abstract class ExcelConditionalFormattingRule : IExcelConditionalFormattingRule
    {
        //Deprecated
        public XmlNode Node { get; }

        public eExcelConditionalFormattingRuleType Type { get; set; }
        public virtual ExcelAddress Address { get; set; }
        public int Priority { get; set; } = 1;
        public bool StopIfTrue { get; set; }
        public bool PivotTable { get; set; }

        ExcelDxfStyleConditionalFormatting _style = null;

        /// <summary>
        /// The style
        /// </summary>
        public ExcelDxfStyleConditionalFormatting Style
        {
            get
            {
                if (_style == null)
                {
                    _style = new ExcelDxfStyleConditionalFormatting(_ws.NameSpaceManager, null, _ws.Workbook.Styles, null);
                }
                return _style;
            }
        }
        //public ExcelDxfStyleConditionalFormatting Style { get; set; }

        internal UInt16 _stdDev = 0;

        //0 is not allowed and will be converted to 1
        public UInt16 StdDev
        {
            get
            {
                return _stdDev;
            }
            set
            {
                _stdDev = value == 0 ? (UInt16)1 : value;
            }
        }

        internal UInt16 _rank = 0;

        /// <summary>
        /// Rank (zero is not allowed and will be converted to 1)
        /// </summary>
        public UInt16 Rank
        {
            get
            {
                return _rank;
            }
            set
            {
                _rank = value == 0 ? (UInt16)1 : value;
            }
        }

        internal string _text = null;

        protected ExcelWorksheet _ws;

        private int _dxfId = -1;

        /// <summary>
        /// The DxfId (Differential Formatting style id)
        /// </summary>
        internal int DxfId
        {
            get
            {
                return _dxfId;
            }
            set
            { _dxfId = value; }
        }


        internal bool IsIconSet 
        {
            get
            {
                return Type == eExcelConditionalFormattingRuleType.ThreeIconSet || Type == eExcelConditionalFormattingRuleType.FourIconSet || Type == eExcelConditionalFormattingRuleType.FiveIconSet;
            }
        }

        internal string _uid = null;

        internal virtual string Uid
        { 
            get
            {
                if (_uid == null)
                {
                    return "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
                }

                return _uid;
            }
            set
            {
                _uid = value;
            }
            
          
        }

        bool _isExtLst = false;

        internal virtual bool IsExtLst 
        { 
            get 
            {
                //Only databars, iconsets and anything with custom formulas can be extLst
                if (Type == eExcelConditionalFormattingRuleType.DataBar)
                {
                    return true;
                }

                if(ExcelAddressBase.RefersToOtherWorksheet(Formula, _ws.Name) || ExcelAddressBase.RefersToOtherWorksheet(Formula2, _ws.Name))
                {
                    return true;
                }

                return _isExtLst;
            }
        }

        #region Constructors
        /// <summary> 
        /// Initalize <see cref="ExcelConditionalFormattingRule"/> from file
        /// </summary>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingRule(eExcelConditionalFormattingRuleType type, ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
        {
            _ws = ws;

            Address = address;
            if(Address == null)
            {
                _isExtLst = true;
            }

            Priority = int.Parse(xr.GetAttribute("priority"));

            Type = type;

            if(!string.IsNullOrEmpty(xr.GetAttribute("id")))
            {
                Uid = xr.GetAttribute("id");
            }

            // Type = (eExcelConditionalFormattingRuleType)Enum.Parse(typeof(eExcelConditionalFormattingRuleType), xr.GetAttribute("type"));

            if (!string.IsNullOrEmpty(xr.GetAttribute("dxfId")))
            {
                DxfId = int.Parse(xr.GetAttribute("dxfId"));
            }

            if(!string.IsNullOrEmpty(xr.GetAttribute("text")))
            {
                _text = xr.GetAttribute("text");
            }

            string timePeriodString = xr.GetAttribute("timePeriod");

            if(!string.IsNullOrEmpty(timePeriodString))
            {
                TimePeriod = timePeriodString.ToEnum<eExcelConditionalFormattingTimePeriodType>();
            }

            ReadClassSpecificXmlNodes(xr);

            xr.Read();

            if (xr.LocalName == "formula" || xr.LocalName == "f")
            {
                _formula = xr.ReadString();
                xr.Read();

                if (xr.LocalName == "formula" || xr.LocalName == "f")
                {
                    _formula2 = xr.ReadString();
                    xr.Read();
                }
            }

            if (address == null)
            {
                if(xr.LocalName == "dxf")
                {
                    ReadExtDxf(xr);
                }
            }

            if (DxfId >= 0 && DxfId < _ws.Workbook.Styles.Dxfs.Count)
            {
                _ws.Workbook.Styles.Dxfs[DxfId].AllowChange = true;  //This Id is referenced by CF, so we can use it when we save.
                _style = ((ExcelDxfStyleBase)_ws.Workbook.Styles.Dxfs[DxfId]).ToDxfConditionalFormattingStyle();    //Clone, so it can be altered without affecting other dxf styles
            }

            var tempAddress = "";

            if (address == null && xr.ReadUntil("cfRule", "sqref", "conditionalFormatting", "extLst") && xr.NodeType == XmlNodeType.EndElement)
            {
                xr.Read();

                if (xr.LocalName == "sqref")
                {
                    tempAddress = xr.ReadString();
                    if (tempAddress == null)
                    {
                        throw new NullReferenceException($"Unable to locate ExtList adress for DataValidation with uid:{Uid}");
                    }
                }
            }
            else
            {
                if (address == null && xr.LocalName == "cfRule" && xr.NodeType == XmlNodeType.EndElement)
                {
                    xr.Read();
                }

                if (xr.LocalName == "sqref")
                {
                    tempAddress = xr.ReadString();
                    if (tempAddress == null)
                    {
                        throw new NullReferenceException($"Unable to locate ExtList adress for DataValidation with uid:{Uid}");
                    }
                }
            }

            if(!string.IsNullOrEmpty(tempAddress))
            {
                Address = new ExcelAddress(tempAddress);
            }
        }

        void ReadExtDxf(XmlReader xr)
        {
            xr.Read();

            if (xr.LocalName == "font")
            {
                xr.Read();
                if(xr.LocalName == "b")
                {
                    Style.Font.Bold = ParseXMlBoolValue(xr);
                    xr.Read();
                }

                if(xr.LocalName == "i")
                {
                    Style.Font.Italic = ParseXMlBoolValue(xr);
                    xr.Read();
                }

                if(xr.LocalName == "strike")
                {
                    Style.Font.Strike = ParseXMlBoolValue(xr);
                    xr.Read();
                }

                if(xr.LocalName == "u")
                {
                    if(xr.GetAttribute("val") == "double")
                    {
                        Style.Font.Underline = ExcelUnderLineType.Double;
                    }
                    else
                    {
                        Style.Font.Underline = ExcelUnderLineType.Single;
                    }
                    xr.Read();
                }

                if(xr.LocalName == "color")
                {
                    ParseColor(Style.Font.Color, xr);
                }
            }


            if (xr.LocalName == "numFmt")
            {
                Style.NumberFormat.NumFmtID = int.Parse(xr.GetAttribute("numFmtId"));
                Style.NumberFormat.Format = xr.GetAttribute("formatCode");
                xr.Read();
            }

            if (xr.LocalName == "fill")
            {
                xr.Read();
                if (xr.LocalName == "patternFill")
                {
                    Style.Fill.Style = eDxfFillStyle.PatternFill;
                    string type = xr.GetAttribute("patternType");
                    Style.Fill.PatternType = string.IsNullOrEmpty(type) ?
                        ExcelFillStyle.None : type.ToEnum<ExcelFillStyle>();
                    xr.Read();

                    if (xr.LocalName == "fgColor")
                    {
                        ParseColor(Style.Fill.PatternColor, xr);
                        xr.Read();
                    }

                    if (xr.LocalName == "bgColor")
                    {
                        ParseColor(Style.Fill.BackgroundColor, xr);

                        if(!string.IsNullOrEmpty(xr.GetAttribute("tint")))
                        {
                            Style.Fill.BackgroundColor.Tint = double.Parse(xr.GetAttribute("tint"));
                        }
                        
                        xr.Read();
                    }
                }
                else
                {
                    Style.Fill.Style = eDxfFillStyle.GradientFill;
                    string degree = xr.GetAttribute("degree");
                    Style.Fill.Gradient.Degree = string.IsNullOrEmpty(degree) ?
                        null : (double?)double.Parse(degree);

                    if(!string.IsNullOrEmpty(xr.GetAttribute("type")))
                    {
                        Style.Fill.Gradient.GradientType = xr.GetAttribute("type").ToEnum<eDxfGradientFillType>();
                    }

                    var doubleString = xr.GetAttribute("left");

                    if (!string.IsNullOrEmpty(doubleString))
                    {
                        Style.Fill.Gradient.Left = double.Parse(doubleString, CultureInfo.InvariantCulture);
                    }

                    if (!string.IsNullOrEmpty(xr.GetAttribute("right")))
                    {
                        Style.Fill.Gradient.Right = double.Parse(xr.GetAttribute("right"), CultureInfo.InvariantCulture);
                    }

                    if (!string.IsNullOrEmpty(xr.GetAttribute("top")))
                    {
                        Style.Fill.Gradient.Top = double.Parse(xr.GetAttribute("top"), CultureInfo.InvariantCulture);
                    }

                    if (!string.IsNullOrEmpty(xr.GetAttribute("bottom")))
                    {
                        Style.Fill.Gradient.Bottom = double.Parse(xr.GetAttribute("bottom"), CultureInfo.InvariantCulture);
                    }

                    xr.Read();
                    ParseColor(Style.Fill.Gradient.Colors.Add(0).Color, xr);
                    xr.Read();

                    ParseColor(Style.Fill.Gradient.Colors.Add(1).Color, xr);
                    xr.Read();
                }
            }

            if (xr.Name == "border")
            {
                do
                {
                    xr.Read();

                    var name = xr.Name;

                    if (name == "left")
                    {
                        Style.Border.Left.Style = xr.GetAttribute("style").ToEnum<ExcelBorderStyle>();
                        ParseColor(Style.Border.Left.Color, xr);
                    }
                    if (name == "right")
                    {
                        Style.Border.Right.Style = xr.GetAttribute("style").ToEnum<ExcelBorderStyle>();
                        ParseColor(Style.Border.Right.Color, xr);
                    }
                    if (name == "top")
                    {
                        Style.Border.Top.Style = xr.GetAttribute("style").ToEnum<ExcelBorderStyle>();
                        ParseColor(Style.Border.Top.Color, xr);
                    }
                    if (name == "bottom")
                    {
                        Style.Border.Bottom.Style = xr.GetAttribute("style").ToEnum<ExcelBorderStyle>();
                        ParseColor(Style.Border.Bottom.Color, xr);
                    }

                } while (xr.Name != "border");
            }
        }

        void ParseColor(ExcelDxfColor col, XmlReader xr)
        {
            xr.Read();
            if (xr.Name == "color" || xr.Name == "bgColor" || xr.Name == "fgColor")
            {
                if (xr.GetAttribute("theme") != null)
                {
                    col.Theme = (Drawing.eThemeSchemeColor)int.Parse(xr.GetAttribute("theme"));
                }
                else if (xr.GetAttribute("rgb") != null)
                {
                    col.Color = ExcelConditionalFormattingHelper.
                        ConvertFromColorCode(xr.GetAttribute("rgb"));
                }
                else if (xr.GetAttribute("auto") != null)
                {
                    col.Auto = xr.GetAttribute("auto") == "1" ? true : false;
                }

                if(!string.IsNullOrEmpty(xr.GetAttribute("tint")))
                {
                    col.Tint = double.Parse(xr.GetAttribute("tint"), CultureInfo.InvariantCulture);
                }
                xr.Read();
            }
        }

        bool ParseXMlBoolValue(XmlReader xr)
        {
            string val = xr.GetAttribute("val");
            if (!string.IsNullOrEmpty(val) && val != "1")
            {
                return false;
            }
            else
            {
                return true;
            }
        }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="original"></param>
            protected ExcelConditionalFormattingRule(ExcelConditionalFormattingRule original)
        {
            _ws = original._ws;
            Rank = original.Rank;
            _formula = original.Formula;
            _formula2 = original.Formula2;
            Operator = original.Operator;
            Type = original.Type;
            PivotTable = original.PivotTable;
            _text = original._text;
            StdDev = original.StdDev;
            DxfId = original.DxfId;
            Address = original.Address;

            if (DxfId >= 0 && DxfId < _ws.Workbook.Styles.Dxfs.Count)
            {
                _ws.Workbook.Styles.Dxfs[DxfId].AllowChange = true;  //This Id is referenced by CF, so we can use it when we save.
                _style = _ws.Workbook.Styles.Dxfs[DxfId].ToDxfConditionalFormattingStyle();    //Clone, so it can be altered without affecting other dxf styles
            }
        }

        internal virtual void ReadClassSpecificXmlNodes(XmlReader xr)
        {

        }

        /// <summary>
        /// Initalize <see cref="ExcelConditionalFormattingRule"/> from variables
        /// </summary>
        /// <param name="type"></param>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingRule(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet)
        {
            FormulaParsing.Utilities.Require.That(worksheet).IsNotNull();

            _ws = worksheet;

            //string.Format()
            //move writing of root node.

            Address = address;
            Priority = priority;
            Type = type;

            if (DxfId >= 0 && DxfId < worksheet.Workbook.Styles.Dxfs.Count)
            {
                worksheet.Workbook.Styles.Dxfs[DxfId].AllowChange = true;  //This Id is referenced by CF, so we can use it when we save.
                _style = ((ExcelDxfStyleBase)worksheet.Workbook.Styles.Dxfs[DxfId]).ToDxfConditionalFormattingStyle();    //Clone, so it can be altered without affecting other dxf styles
            }
        }
        #endregion Constructors

        /// <summary>
        /// Above average
        /// In Excel: Default:True, use=optional
        /// </summary>
        internal protected bool? AboveAverage
        {
            get
            {
                return (Type == eExcelConditionalFormattingRuleType.BelowAverage)
                  || (Type == eExcelConditionalFormattingRuleType.BelowOrEqualAverage)
                  || (Type == eExcelConditionalFormattingRuleType.BelowStdDev)
                  ? false : true;
            }
        }

        /// <summary>
        /// EqualAverage
        /// </summary>
        internal protected bool? EqualAverage
        {
            get
            {
                // Equal Avarege only if TRUE
                return (Type == eExcelConditionalFormattingRuleType.AboveOrEqualAverage)
                  || (Type == eExcelConditionalFormattingRuleType.BelowOrEqualAverage)
                  ? true : false;
            }
        }

        /// <summary>
        /// Bottom attribute
        /// </summary>
        internal protected bool? Bottom
        {
            get
            {
                return (Type == eExcelConditionalFormattingRuleType.Bottom)
                  || (Type == eExcelConditionalFormattingRuleType.BottomPercent) 
                  ? true : false;
            }
        }

        /// <summary>
        /// Percent attribute
        /// </summary>
        internal protected bool? Percent
        {
            get
            {
                return ((Type == eExcelConditionalFormattingRuleType.BottomPercent)
                  || (Type == eExcelConditionalFormattingRuleType.TopPercent))
                  ? true : false;
            }
        }

        /// <summary>
        /// TimePeriod
        /// </summary>
        internal protected eExcelConditionalFormattingTimePeriodType? TimePeriod { get; set; } = null;

        /// <summary>
        /// Operator
        /// </summary>
        internal protected eExcelConditionalFormattingOperatorType? Operator { get; set; } = null;

        internal string _formula;
        internal string _formula2;

        /// <summary>
        /// Formula
        /// </summary>
        public virtual string Formula 
        { 
            get { return _formula; } 
            set
            {
                _formula = value; 
            } 
        }

        /// <summary>
        /// Formula2
        /// Note, no longer Requires Formula to be set before it.
        /// But will still throw error if both formulas not filled at save time.
        /// </summary>
        public virtual string Formula2
        {
            get { return _formula2; }
            set 
            {
                _formula2 = value;
            }
        }
        private ExcelConditionalFormattingAsType _as = null;
        /// <summary>
        /// Provides access to type conversion for all conditional formatting rules.
        /// </summary>
        public ExcelConditionalFormattingAsType As
        {
            get
            {
                if (_as == null)
                {
                    _as = new ExcelConditionalFormattingAsType(this);
                }
                return _as;
            }
        }

        public void SetStyle(ExcelDxfStyleConditionalFormatting style)
        {
            _style = style;
        }

        internal string GetAttributeType()
        {
            return ExcelConditionalFormattingRuleType.GetAttributeByType(Type);
        }

        internal abstract ExcelConditionalFormattingRule Clone();
    }
}
