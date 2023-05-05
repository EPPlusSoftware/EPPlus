﻿using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Xml;
using OfficeOpenXml.Utils.Extensions;

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

        internal string Text = null;

        private ExcelWorksheet _ws;

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

        internal virtual bool IsExtLst 
        { 
            get 
            {
                //Only databars, iconsets and anything with custom formulas can be extLst
                if (Type == eExcelConditionalFormattingRuleType.DataBar)
                {
                    return true;
                }

                return false;
            } 
        }

        #region Constructors
        /// <summary> 
        /// Initalize <see cref="ExcelConditionalFormattingRule"/> from file
        /// </summary>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingRule(eExcelConditionalFormattingRuleType type, ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
        {
            Address = address;

            Priority = int.Parse(xr.GetAttribute("priority"));

            Type = type;

            // Type = (eExcelConditionalFormattingRuleType)Enum.Parse(typeof(eExcelConditionalFormattingRuleType), xr.GetAttribute("type"));

            if (!string.IsNullOrEmpty(xr.GetAttribute("dxfId")))
            {
                DxfId = int.Parse(xr.GetAttribute("dxfId"));
            }

            if(!string.IsNullOrEmpty(xr.GetAttribute("text")))
            {
                Text = xr.GetAttribute("text");
            }

            string timePeriodString = xr.GetAttribute("timePeriod");

            if(!string.IsNullOrEmpty(timePeriodString))
            {
                TimePeriod = timePeriodString.ToEnum<eExcelConditionalFormattingTimePeriodType>();
            }

            ReadClassSpecificXmlNodes(xr);

            xr.Read();

            if (xr.LocalName == "formula")
            {
                Formula = xr.ReadString();
                xr.Read();

                if (xr.LocalName == "formula")
                {
                    Formula2 = xr.ReadString();
                    xr.Read();
                }
            }

            _ws = ws;

            if (DxfId >= 0 && DxfId < _ws.Workbook.Styles.Dxfs.Count)
            {
                _ws.Workbook.Styles.Dxfs[DxfId].AllowChange = true;  //This Id is referenced by CF, so we can use it when we save.
                _style = ((ExcelDxfStyleBase)_ws.Workbook.Styles.Dxfs[DxfId]).ToDxfConditionalFormattingStyle();    //Clone, so it can be altered without affecting other dxf styles
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
            Formula = original.Formula;
            Formula2 = original.Formula2;
            Operator = original.Operator;
            Type = original.Type;
            PivotTable = original.PivotTable;
            Text = original.Text;
            StdDev = original.StdDev;
            DxfId = original.DxfId;

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
            FormulaParsing.Utilities.Require.That(address).IsNotNull();
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

        string _formula;
        string _formula2;

        /// <summary>
        /// Formula
        /// </summary>
        public string Formula 
        { 
            get { return _formula; } 
            set { _formula = value; } 
        }

        /// <summary>
        /// Formula2
        /// Note, no longer Requires Formula to be set before it.
        /// But will still throw error if both formulas not filled at save time.
        /// </summary>
        public string Formula2
        {
            get { return _formula2; }
            set { _formula2 = ConvertUtil.ExcelEscapeAndEncodeString(value); }
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
