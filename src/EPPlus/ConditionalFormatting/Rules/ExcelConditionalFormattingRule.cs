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
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Xml;
using OfficeOpenXml.Utils.Extensions;
using OfficeOpenXml.Style;
using System.Globalization;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// Abstract base class for all ConditionalFormattingRules
    /// </summary>
    public abstract class ExcelConditionalFormattingRule : IExcelConditionalFormattingRule
    {
        /// <summary>
        /// The type of conditional formatting rule.
        /// </summary>
        public eExcelConditionalFormattingRuleType Type { get; set; }

        internal ExcelAddress _address = null;

        /// <summary>
        /// <para>The range over which these conditional formatting rules apply.</para>
        /// </summary>
        public virtual ExcelAddress Address
        { get
            {
                return _address;
            }
            set
            {
                _address = value;
                //QuadItem = new QuadItem<ExcelConditionalFormattingRule>(){ this, _address }
                //QuadRange = new QuadRange(value);
            }
        }

        //internal QuadRange QuadRange { get; private set; }
        //internal QuadItem<ExcelConditionalFormattingRule> QuadItem { get; private set; }

        internal int _priority = 1;

        /// <summary>
        /// The priority of the rule. 
        /// 1 is highest priority. 2 second highest etc.
        /// </summary>
        public int Priority
        {
            get
            {
                return _priority;
            }
            set
            {
                _ws.ConditionalFormatting.ChangePriority(this, value);
                _priority = value;
            }
        }
        /// <summary>
        /// If this property is true, no rules with lower priority should be applied over this rule.
        /// </summary>
        public bool StopIfTrue { get; set; }

        bool _pivotTable = false;

        /// <summary>
        /// Indicates that the conditional formatting is associated with a PivotTable
        /// </summary>
        public bool PivotTable
        {
            get { return _pivotTable; }
            set
            {
                if (Address == null)
                {
                    _isExtLst = false;
                }
                _pivotTable = value;
            }
        }

        internal ExcelDxfStyleConditionalFormatting _style = null;

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

        internal UInt16 _stdDev = 0;

        /// <summary>
        /// 0 is not allowed and will be converted to 1
        /// </summary>
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

        /// <summary>
        /// Internal worksheet reference
        /// </summary>
        internal ExcelWorksheet _ws;

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
                    _uid = Guid.NewGuid().ToString().ToUpperInvariant();
                    return _uid;
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

                if (ExcelAddressBase.RefersToOtherWorksheet(Formula, _ws.Name) || ExcelAddressBase.RefersToOtherWorksheet(Formula2, _ws.Name))
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
        /// <param name="type"></param>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        internal ExcelConditionalFormattingRule(eExcelConditionalFormattingRuleType type, ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
        {
            _ws = ws;

            Address = address;

            if (Address == null && !PivotTable)
            {
                _isExtLst = true;
            }

            _priority = int.Parse(xr.GetAttribute("priority"));

            Type = type;

            if (!string.IsNullOrEmpty(xr.GetAttribute("stopIfTrue")))
            {
                StopIfTrue = int.Parse(xr.GetAttribute("stopIfTrue")) == 1;
            }


            if (!string.IsNullOrEmpty(xr.GetAttribute("id")))
            {
                Uid = xr.GetAttribute("id");
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("dxfId")))
            {
                DxfId = int.Parse(xr.GetAttribute("dxfId"));
            }

            if (!string.IsNullOrEmpty(xr.GetAttribute("text")))
            {
                _text = xr.GetAttribute("text");
            }

            string timePeriodString = xr.GetAttribute("timePeriod");

            if (!string.IsNullOrEmpty(timePeriodString))
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
                if (xr.LocalName == "dxf")
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

            if (!string.IsNullOrEmpty(tempAddress))
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
                if (xr.LocalName == "b")
                {
                    Style.Font.Bold = ParseXMlBoolValue(xr);
                    xr.Read();
                }

                if (xr.LocalName == "i")
                {
                    Style.Font.Italic = ParseXMlBoolValue(xr);
                    xr.Read();
                }

                if (xr.LocalName == "strike")
                {
                    Style.Font.Strike = ParseXMlBoolValue(xr);
                    xr.Read();
                }

                if (xr.LocalName == "u")
                {
                    if (xr.GetAttribute("val") == "double")
                    {
                        Style.Font.Underline = ExcelUnderLineType.Double;
                    }
                    else
                    {
                        Style.Font.Underline = ExcelUnderLineType.Single;
                    }
                    xr.Read();
                }

                if (xr.LocalName == "color")
                {
                    ParseColor(Style.Font.Color, xr);
                }

                xr.Read();
            }


            if (xr.LocalName == "numFmt")
            {
                Style.NumberFormat.Format = xr.GetAttribute("formatCode");
                Style.NumberFormat.NumFmtID = int.Parse(xr.GetAttribute("numFmtId"));
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
                        null : type.ToEnum<ExcelFillStyle>();
                    xr.Read();

                    if (xr.LocalName == "fgColor")
                    {
                        ParseColor(Style.Fill.PatternColor, xr);
                    }

                    if (xr.LocalName == "bgColor")
                    {
                        ParseColor(Style.Fill.BackgroundColor, xr);
                    }

                    xr.Read();
                    xr.Read();
                }
                else
                {
                    Style.Fill.Style = eDxfFillStyle.GradientFill;
                    string degree = xr.GetAttribute("degree");
                    Style.Fill.Gradient.Degree = string.IsNullOrEmpty(degree) ?
                        null : (double?)double.Parse(degree, CultureInfo.InvariantCulture);

                    if (!string.IsNullOrEmpty(xr.GetAttribute("type")))
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
                    xr.Read();
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
            if (xr.LocalName != "color" && xr.LocalName != "bgColor" && xr.LocalName != "fgColor")
            {
                xr.Read();
            }

            if (xr.LocalName == "color" || xr.LocalName == "bgColor" || xr.LocalName == "fgColor")
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

                if (!string.IsNullOrEmpty(xr.GetAttribute("tint")))
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
        /// Copy constructor
        /// </summary>
        /// <param name="original"></param>
        /// <param name="newWorksheet">In case cloning from another worksheet</param>
        protected ExcelConditionalFormattingRule(ExcelConditionalFormattingRule original, ExcelWorksheet newWorksheet = null)
        {
            if (newWorksheet == null)
            {
                _ws = original._ws;
            }
            else
            {
                _ws = newWorksheet;
            }

            if (original.Rank != 0)
            {
                Rank = original.Rank;
            }

            _formula = original.Formula;
            _formula2 = original.Formula2;
            Operator = original.Operator;
            Type = original.Type;
            PivotTable = original.PivotTable;
            _text = original._text;
            StopIfTrue = original.StopIfTrue;

            if (original._stdDev != 0)
            {
                StdDev = original.StdDev;
            }

            DxfId = original.DxfId;
            Address = original.Address;
            Priority = original.Priority;

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
            _priority = priority;
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

        internal string calculatedFormula1 = null;
        internal string calculatedFormula2 = null;

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

        internal virtual bool ShouldApplyToCell(ExcelAddress address)
        {
            if (address != null) { return true; }
            return false;
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

        internal void SetStyle(ExcelDxfStyleConditionalFormatting style)
        {
            _style = style;
        }

        internal string GetAttributeType()
        {
            return ExcelConditionalFormattingRuleType.GetAttributeByType(Type);
        }

        internal abstract ExcelConditionalFormattingRule Clone(ExcelWorksheet sheet = null);

        internal virtual void RemoveTempExportData()
        {
            calculatedFormula1 = null;
            calculatedFormula2 = null;
        }

        internal static ISourceCodeTokenizer _tokenizer = SourceCodeTokenizer.Default;
        internal IList<Token> Tokens;

        internal void SetTokens(string worksheet, string aFormula)
        {
            if (Tokens == null)
            {
                if(aFormula != null)
                {
                    Tokens = _tokenizer.Tokenize(aFormula, worksheet);
                }
            }
        }

        internal string GetCellFormula(ExcelAddress address, bool getFormula2 = false)
        {
            return GetFormula(address._fromRow, address._fromCol, getFormula2);
        }

        internal string GetFormula(int row, int column, bool isFormula2 = false)
        {
            var relevantFormula = isFormula2 == true ? Formula2 : Formula;
            if (Address.Start.Row == row && Address.Start.Column == column)
            {
                if(relevantFormula == null)
                {
                    return "";
                }
                return relevantFormula;
            }

            SetTokens(_ws.Name, relevantFormula);

            string f = "";
            if(Tokens == null)
            {
                return "";
            }
            for (int i = 0; i < Tokens.Count; i++)
            {
                var token = Tokens[i];
                if (token.TokenTypeIsSet(TokenType.CellAddress))
                {
                    var a = new ExcelFormulaAddress(token.Value, (ExcelWorksheet)null);
                    if (a.IsFullColumn)
                    {
                        if (a.IsFullRow)
                        {
                            f += token.Value;
                        }
                        else
                        {
                            f += a.GetOffset(0, column - Address.Start.Column, true);
                        }
                    }
                    else if (a.IsFullRow)
                    {
                        f += a.GetOffset(row - Address.Start.Row, 0, true);
                    }
                    else
                    {
                        if (a.Table != null)
                        {
                            f += token.Value;
                        }
                        else
                        {
                            f += a.GetOffset(row - Address.Start.Row, column - Address.Start.Column, true);
                        }
                    }
                }
                else if (token.TokenTypeIsSet(TokenType.FullRowAddress))
                {
                    if (token.Value.StartsWith("$") == false)
                    {
                        if (int.TryParse(token.Value, out int r))
                        {
                            r += row - Address.Start.Row;
                            if (r >= 1 && r <= ExcelPackage.MaxRows)
                            {
                                f += r.ToString(CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                f += "#REF!";
                            }
                        }
                        else
                        {
                            f += "#REF!";
                        }
                    }
                    else
                    {
                        f += token.Value;
                    }
                }
                else if (token.TokenTypeIsSet(TokenType.FullColumnAddress))
                {
                    if (token.Value.StartsWith("$") == false)
                    {
                        var c = ExcelCellBase.GetColumn(token.Value);
                        c += column - Address.Start.Column;
                        if (c >= 1 && c <= ExcelPackage.MaxColumns)
                        {
                            f += ExcelCellBase.GetColumnLetter(c);
                        }
                        else
                        {
                            f += "#REF!";
                        }
                    }
                    else
                    {
                        f += token.Value;
                    }
                }
                else
                {
                    f += token.Value;
                }

            }
            Tokens = null;
            return f;
        }
    }
}
