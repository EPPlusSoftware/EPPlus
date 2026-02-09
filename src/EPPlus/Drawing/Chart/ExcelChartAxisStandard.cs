/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/22/2020         EPPlus Software AB       Added this class
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.EnumUtils;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;
using static OfficeOpenXml.Style.XmlAccess.ExcelNumberFormatXml;
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// An axis for a standard chart. 
    /// </summary>
    public sealed class ExcelChartAxisStandard : ExcelChartAxis
    {
        internal readonly static string[] _schemaNodeOrderDateShared = new string[]
        {
            "axId",
            "scaling",
            "delete",
            "axPos",
            "majorGridlines",
            "minorGridlines",
            "title",
            "numFmt",
            "majorTickMark",
            "minorTickMark",
            "tickLblPos",
            "spPr",
            "txPr",
            "crossAx",
            "crosses",
            "crossesAt"
        };
        internal static string[] _schemaNodeOrderCat = new string[] { "auto", "lblAlgn", "lblOffset", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "extLst" };
        internal static string[] _schemaNodeOrderDate = new string[] { "auto", "lblOffset", "baseTimeUnit", "majorUnit", "majorTimeUnit", "minorUnit", "minorTimeUnit", "extLst" };
        internal static string[] _schemaNodeOrderSer = new string[] { "tickLblSkip", "tickMarkSkip", "extLst" };
        internal static string[] _schemaNodeOrderVal = new string[] { "crossBetween", "majorUnit", "minorUnit", "dispUnits", "extLst" };

        internal ExcelChartAxisStandard(ExcelChart chart, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string nsPrefix, int index) : base(chart, nameSpaceManager, topNode, nsPrefix)
        {
            AddSchemaNodeOrder(new string[] { "axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines", "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses", "crossesAt", "crossBetween", "auto", "lblOffset", "baseTimeUnit", "majorUnit", "majorTimeUnit", "minorUnit", "minorTimeUnit", "tickLblSkip", "tickMarkSkip", "dispUnits", "noMultiLvlLbl", "logBase", "orientation", "max", "min" },
                ExcelDrawing._schemaNodeOrderSpPr);

            Index = index;
        }
        internal override string Id
        {
            get
            {
                return GetXmlNodeString("c:axId/@val");
            }
        }

        const string _majorTickMark = "c:majorTickMark/@val";
        /// <summary>
        /// Get or Sets the major tick marks for the axis. 
        /// </summary>
        public override eAxisTickMark MajorTickMark
        {
            get
            {
                var v = GetXmlNodeString(_majorTickMark);
                if (string.IsNullOrEmpty(v))
                {
                    return eAxisTickMark.Cross;
                }
                else
                {
                    try
                    {
                        return (eAxisTickMark)Enum.Parse(typeof(eAxisTickMark), v, true);
                    }
                    catch
                    {
                        return eAxisTickMark.Cross;
                    }
                }
            }
            set
            {
                SetXmlNodeString(_majorTickMark, value.ToString().ToLower(CultureInfo.InvariantCulture));
            }
        }
        const string _minorTickMark = "c:minorTickMark/@val";
        /// <summary>
        /// Get or Sets the minor tick marks for the axis. 
        /// </summary>
        public override eAxisTickMark MinorTickMark
        {
            get
            {
                var v = GetXmlNodeString(_minorTickMark);
                if (string.IsNullOrEmpty(v))
                {
                    return eAxisTickMark.Cross;
                }
                else
                {
                    try
                    {
                        return (eAxisTickMark)Enum.Parse(typeof(eAxisTickMark), v, true);
                    }
                    catch
                    {
                        return eAxisTickMark.Cross;
                    }
                }
            }
            set
            {
                SetXmlNodeString(_minorTickMark, value.ToString().ToLower(CultureInfo.InvariantCulture));
            }
        }
        private string AXIS_POSITION_PATH = "c:axPos/@val";
        /// <summary>
        /// Where the axis is located
        /// </summary>
        public override eAxisPosition AxisPosition
        {
            get
            {
                switch (GetXmlNodeString(AXIS_POSITION_PATH))
                {
                    case "b":
                        return eAxisPosition.Bottom;
                    case "r":
                        return eAxisPosition.Right;
                    case "t":
                        return eAxisPosition.Top;
                    default:
                        return eAxisPosition.Left;
                }
            }
            internal set
            {
                SetXmlNodeString(AXIS_POSITION_PATH, value.ToString().ToLower(CultureInfo.InvariantCulture).Substring(0, 1));
            }
        }
        /// <summary>
        /// Chart axis title
        /// </summary>
        public new ExcelChartTitleStandard Title
        {
            get
            {
                return (ExcelChartTitleStandard)GetTitle();
            }
        }
        /// <summary>
        /// True if the axis has a title.
        /// </summary>
        public override bool HasTitle
        {
            get
            {
                return TopNode.SelectSingleNode("c:title", NameSpaceManager) != null;
            }
        }
        internal override ExcelChartTitle GetTitle()
        {
            if (_title == null)
            {
                AddTitleNode();
                _title = new ExcelChartTitleStandard(_chart, NameSpaceManager, TopNode, "c");
            }
            return _title;
        }
        const string _minValuePath = "c:scaling/c:min/@val";
        /// <summary>
        /// Minimum value for the axis.
        /// Null is automatic
        /// </summary>
        public override double? MinValue
        {
            get
            {
                return GetXmlNodeDoubleNull(_minValuePath);
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_minValuePath, true);
                }
                else
                {
                    SetXmlNodeString(_minValuePath, ((double)value).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        const string _maxValuePath = "c:scaling/c:max/@val";
        /// <summary>
        /// Max value for the axis.
        /// Null is automatic
        /// </summary>
        public override double? MaxValue
        {
            get
            {
                return GetXmlNodeDoubleNull(_maxValuePath);
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_maxValuePath, true);
                }
                else
                {
                    SetXmlNodeString(_maxValuePath, ((double)value).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        const string _lblPos = "c:tickLblPos/@val";
        /// <summary>
        /// The Position of the labels
        /// </summary>
        public override eTickLabelPosition LabelPosition
        {
            get
            {
                var v = GetXmlNodeString(_lblPos);
                if (string.IsNullOrEmpty(v))
                {
                    return eTickLabelPosition.NextTo;
                }
                else
                {
                    try
                    {
                        return (eTickLabelPosition)Enum.Parse(typeof(eTickLabelPosition), v, true);
                    }
                    catch
                    {
                        return eTickLabelPosition.NextTo;
                    }
                }
            }
            set
            {
                string lp = value.ToString();
                SetXmlNodeString(_lblPos, lp.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + lp.Substring(1, lp.Length - 1));
            }
        }
        const string _crossesPath = "c:crosses/@val";
        /// <summary>
        /// Where the axis crosses
        /// </summary>
        public override eCrosses Crosses
        {
            get
            {
                var v = GetXmlNodeString(_crossesPath);
                if (string.IsNullOrEmpty(v))
                {
                    return eCrosses.AutoZero;
                }
                else
                {
                    try
                    {
                        return (eCrosses)Enum.Parse(typeof(eCrosses), v, true);
                    }
                    catch
                    {
                        return eCrosses.AutoZero;
                    }
                }
            }
            set
            {
                var v = value.ToString();
                v = v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1, v.Length - 1);
                SetXmlNodeString(_crossesPath, v);
            }

        }
        const string _crossBetweenPath = "c:crossBetween/@val";
        /// <summary>
        /// How the axis are crossed
        /// </summary>
        public override eCrossBetween CrossBetween
        {
            get
            {
                var v = GetXmlNodeString(_crossBetweenPath);
                if (string.IsNullOrEmpty(v))
                {
                    return eCrossBetween.Between;
                }
                else
                {
                    try
                    {
                        return (eCrossBetween)Enum.Parse(typeof(eCrossBetween), v, true);
                    }
                    catch
                    {
                        return eCrossBetween.Between;
                    }
                }
            }
            set
            {
                var v = value.ToString();
                v = v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1);
                SetXmlNodeString(_crossBetweenPath, v);
            }
        }
        const string _crossesAtPath = "c:crossesAt/@val";
        /// <summary>
        /// The value where the axis cross. 
        /// Null is automatic
        /// </summary>
        public override double? CrossesAt
        {
            get
            {
                return GetXmlNodeDoubleNull(_crossesAtPath);
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_crossesAtPath, true);
                }
                else
                {
                    SetXmlNodeString(_crossesAtPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
        /// <summary>
        /// If the axis is deleted
        /// </summary>
        public override bool Deleted
        {
            get
            {
                return GetXmlNodeBool("c:delete/@val");
            }
            set
            {
                SetXmlNodeBool("c:delete/@val", value);
            }
        }
        const string _ticLblPos_Path = "c:tickLblPos/@val";
        /// <summary>
        /// Position of the Lables
        /// </summary>
        public override eTickLabelPosition TickLabelPosition
        {
            get
            {
                string v = GetXmlNodeString(_ticLblPos_Path);
                if (v == "")
                {
                    return eTickLabelPosition.None;
                }
                else
                {
                    return (eTickLabelPosition)Enum.Parse(typeof(eTickLabelPosition), v, true);
                }
            }
            set
            {
                string v = value.ToString();
                v = v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1, v.Length - 1);
                SetXmlNodeString(_ticLblPos_Path, v);
            }
        }
        const string _displayUnitPath = "c:dispUnits/c:builtInUnit/@val";
        const string _custUnitPath = "c:dispUnits/c:custUnit/@val";
        /// <summary>
        /// The scaling value of the display units for the value axis
        /// </summary>
        public override double DisplayUnit
        {
            get
            {
                string v = GetXmlNodeString(_displayUnitPath);
                if (string.IsNullOrEmpty(v))
                {
                    var c = GetXmlNodeDoubleNull(_custUnitPath);
                    if (c == null)
                    {
                        return 0;
                    }
                    else
                    {
                        return c.Value;
                    }
                }
                else
                {
                    try
                    {
                        return (double)(long)Enum.Parse(typeof(eBuildInUnits), v, true);
                    }
                    catch
                    {
                        return 0;
                    }
                }
            }
            set
            {
                if (AxisType == eAxisType.Val && value >= 0)
                {
                    foreach (var v in Enum.GetValues(typeof(eBuildInUnits)))
                    {
                        if ((double)(long)v == value)
                        {
                            DeleteNode(_custUnitPath, true);
                            SetXmlNodeString(_displayUnitPath, ((eBuildInUnits)value).ToString());
                            return;
                        }
                    }
                    DeleteNode(_displayUnitPath, true);
                    if (value != 0)
                    {
                        SetXmlNodeString(_custUnitPath, value.ToString(CultureInfo.InvariantCulture));
                    }
                }
            }
        }
        const string _majorUnitPath = "c:majorUnit/@val";
        const string _majorUnitCatPath = "c:tickLblSkip/@val";
        /// <summary>
        /// Major unit for the axis.
        /// Null is automatic
        /// </summary>
        public override double? MajorUnit
        {
            get
            {
                if (AxisType == eAxisType.Cat)
                {
                    return GetXmlNodeDoubleNull(_majorUnitCatPath);
                }
                else
                {
                    return GetXmlNodeDoubleNull(_majorUnitPath);
                }
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_majorUnitPath, true);
                    DeleteNode(_majorUnitCatPath, true);
                }
                else
                {
                    if (AxisType == eAxisType.Cat)
                    {
                        SetXmlNodeString(_majorUnitCatPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        SetXmlNodeString(_majorUnitPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                    }
                }
            }
        }
        const string _majorTimeUnitPath = "c:majorTimeUnit/@val";
        /// <summary>
        /// Major time unit for the axis.
        /// Null is automatic
        /// </summary>
        public override eTimeUnit? MajorTimeUnit
        {
            get
            {
                var v = GetXmlNodeString(_majorTimeUnitPath);
                if (string.IsNullOrEmpty(v))
                {
                    return null;
                }
                else
                {
                    return v.ToEnum(eTimeUnit.Years);
                }
            }
            set
            {
                if (value.HasValue)
                {
                    SetXmlNodeString(_majorTimeUnitPath, value.ToEnumString());
                }
                else
                {
                    DeleteNode(_majorTimeUnitPath, true);
                }
            }
        }
        const string _minorUnitPath = "c:minorUnit/@val";
        const string _minorUnitCatPath = "c:tickMarkSkip/@val";
        /// <summary>
        /// Minor unit for the axis.
        /// Null is automatic
        /// </summary>
        public override double? MinorUnit
        {
            get
            {
                if (AxisType == eAxisType.Cat)
                {
                    return GetXmlNodeDoubleNull(_minorUnitCatPath);
                }
                else
                {
                    return GetXmlNodeDoubleNull(_minorUnitPath);
                }
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_minorUnitPath, true);
                    DeleteNode(_minorUnitCatPath, true);
                }
                else
                {
                    if (AxisType == eAxisType.Cat)
                    {
                        SetXmlNodeString(_minorUnitCatPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        SetXmlNodeString(_minorUnitPath, ((double)value).ToString(CultureInfo.InvariantCulture));
                    }
                }
            }
        }
        const string _minorTimeUnitPath = "c:minorTimeUnit/@val";
        /// <summary>
        /// Minor time unit for the axis.
        /// Null is automatic
        /// </summary>
        public override eTimeUnit? MinorTimeUnit
        {
            get
            {
                var v = GetXmlNodeString(_minorTimeUnitPath);
                if (string.IsNullOrEmpty(v))
                {
                    return null;
                }
                else
                {
                    return v.ToEnum(eTimeUnit.Years);
                }
            }
            set
            {
                if (value.HasValue)
                {
                    SetXmlNodeString(_minorTimeUnitPath, value.ToEnumString());
                }
                else
                {
                    DeleteNode(_minorTimeUnitPath, true);
                }
            }
        }
        const string _logbasePath = "c:scaling/c:logBase/@val";
        /// <summary>
        /// The base for a logaritmic scale
        /// Null for a normal scale
        /// </summary>
        public override double? LogBase
        {
            get
            {
                return GetXmlNodeDoubleNull(_logbasePath);
            }
            set
            {
                if (value == null)
                {
                    DeleteNode(_logbasePath, true);
                }
                else
                {
                    double v = ((double)value);
                    if (v < 2 || v > 1000)
                    {
                        throw (new ArgumentOutOfRangeException("Value must be between 2 and 1000"));
                    }
                    SetXmlNodeString(_logbasePath, v.ToString("0.0", CultureInfo.InvariantCulture));
                }
            }
        }
        const string _orientationPath = "c:scaling/c:orientation/@val";
        /// <summary>
        /// Axis orientation
        /// </summary>
        public override eAxisOrientation Orientation
        {
            get
            {
                string v = GetXmlNodeString(_orientationPath);
                if (v == "")
                {
                    return eAxisOrientation.MinMax;
                }
                else
                {
                    return (eAxisOrientation)Enum.Parse(typeof(eAxisOrientation), v, true);
                }
            }
            set
            {
                string s = value.ToString();
                s = s.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + s.Substring(1, s.Length - 1);
                SetXmlNodeString(_orientationPath, s);
            }
        }

        internal override eAxisType AxisType
        {
            get
            {
                try
                {
                    var axType = TopNode.LocalName.Substring(0, TopNode.LocalName.Length - 2);
                    if (axType == "ser") return eAxisType.Serie;
                    return (eAxisType)Enum.Parse(typeof(eAxisType), axType, true);
                }
                catch
                {
                    return eAxisType.Val;
                }
            }
        }
        /// <summary>
        /// Adds the axis title and styles it according to the style selected in the StyleManager
        /// </summary>
        /// <param name="linkedCell">The cell linked to the title.</param>
        public void AddTitle(ExcelRangeBase linkedCell)
        {
            Title.LinkedCell = linkedCell;
            _chart.ApplyStyleOnPart(Title, _chart._styleManager?.Style?.AxisTitle);
        }
        /// <summary>
        /// The axis data is specified on the range data, not specified, for a date or category axis
        /// </summary>
        public bool Auto
        {
            get
            {
                return GetXmlNodeBool("c:auto/@val");
            }
            set
            {
                SetXmlNodeBool("c:auto/@val", value);
            }
        }
        ExcelLayout _layout = null;
        /// <summary>
        /// Contains layout properties, if the title is manually positioned.
        /// </summary>
        public ExcelLayout Layout
        {
            get
            {
                if (_layout == null)
                {
                    _layout = new ExcelLayout(NameSpaceManager, TopNode, "c:layout", "", SchemaNodeOrder);
                }
                return _layout;
            }
        }
        /// <summary>
        /// The index for the axis. 
        /// </summary>
        public int Index { get; private set; }
        internal bool IsDate
        {
            get
            {
                if (AxisType == eAxisType.Date)
                {
                    return true;
                }
                else if (AxisType != eAxisType.Cat)
                {
                    var nf = new ExcelFormatTranslator(Format, 0);
                    if (nf.DataType == eFormatType.DateTime)
                    {
                        return true;
                    }
                }
                return false;
            }
        }

        internal bool IsYAxis 
        { 
            get
            {
                foreach(var ct in _chart.PlotArea.ChartTypes)
                {
                    if (ct.YAxis == this)
                    {
                        return true;
                    }
                }
                return false;
            }
        }
        internal bool IsXAxis
        {
            get
            {
                foreach (var ct in _chart.PlotArea.ChartTypes)
                {
                    if (ct.XAxis == this)
                    {
                        return true;
                    }
                }
                return false;
            }
        }
        internal override object[] GetAxisValues(out bool isCount)
        {
            List<List<object>> values;
            GetSeriesValues(out isCount, out values);
            var dl = values.SelectMany(x => x).Distinct().ToList();
            dl.Sort();
            if (Orientation == eAxisOrientation.MaxMin)
            {
                dl.Reverse();
            }
            return dl.ToArray();
        }

        internal void GetSeriesValues(out bool isCount, out List<List<object>> values)
        {
            values = new List<List<object>>();
            isCount = false;
            List<object> pl = new List<object>();
            foreach (var ct in _chart.PlotArea.ChartTypes)
            {
                foreach (var serie in _chart.Series)
                {
                    var l = new List<object>();
                    values.Add(l);
                    if (AxisType == eAxisType.Cat)
                    {
                        if (string.IsNullOrEmpty(serie.XSeries) && (serie.NumberLiteralsX == null || serie.NumberLiteralsX.Length == 0) && (serie.StringLiteralsX == null || serie.StringLiteralsX?.Length == 0))
                        {
                            AddCountFromSeries(l, serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                            isCount = true;
                        }
                        else
                        {
                            AddFromSerie(l, serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX, true);
                        }
                    }
                    else
                    {
                        if (ct.YAxis == this)
                        {
                            AddFromSerie(l, serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY, false, pl);
                        }
                        else if (ct.XAxis == this)
                        {
                            AddFromSerie(l, serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX, false);
                        }
                    }
                    pl = l;
                }
            }
            if (_chart.IsTypePercentStacked() && IsYAxis)
            {
                CalculateStacked100(values);
            }
        }

        internal static void CalculateStacked100(List<List<object>> values)
        {
            for (int i = 0; i < values[0].Count; i++)
            {
                double rowAbsSum = 0;
                for (int k = 0; k < values.Count; k++)
                {
                    var v = values[k][i];
                    if (ConvertUtil.IsExcelNumeric(v))
                    {
                        rowAbsSum += Math.Abs(ConvertUtil.GetValueDouble(v, true, true));
                    }
                }
                var pv = 0D;
                for (int j=0;j< values.Count;j++)
                {
                    var val = values[j][i];
                    if (ConvertUtil.IsExcelNumeric(val))
                    {
                        double d = pv + ConvertUtil.GetValueDouble(val, true, true);
                        if (rowAbsSum != 0)
                        {
                            values[j][i] = d / rowAbsSum;
                        }   
                        else
                        {
                            values[j][i] = 0;
                        }
                        pv = d;
                    }
                }
            }
        }
        private void AddCountFromSeries(List<object> l, string address, double[] numberLiterals, string[] stringLiterals)
        {
            l.Add(1);
            if (numberLiterals?.Length > 0)
            {
                l.Add(numberLiterals.Length);
            }
            else if (stringLiterals?.Length > 0)
            {
                l.Add(stringLiterals.Length);
            }
            else
            {
                var a = new ExcelAddressBase(address);
                l.Add(Math.Max(a.Rows, a.Columns));
            }
        }

        private void AddFromSerie(List<object> list, string address, double[] numberLiterals, string[] stringLiterals, bool useText, List<object> prevList=null)
        {
            var isStacked = _chart.IsTypeStacked() && prevList != null;
            if (numberLiterals?.Length > 0)
            {
                for (int i = 0; i < numberLiterals.Length; i++)
                {
                    var n = numberLiterals[i];
                    if (isStacked)
                    {
                        n += GetSerieValue(prevList, i);
                    }
                    list.Add(n);
                }
            }
            else if (stringLiterals?.Length > 0)
            {
                foreach (var s in stringLiterals)
                {
                    list.Add(s);
                }
            }
            else
            {
                var a = new ExcelAddressBase(address);
                var ws = _chart.WorkSheet.Workbook.Worksheets[a.WorkSheetName];
                if (ws != null)
                {
                    var range = ws.Cells[a.Address];
                    var i = 0;
                    for(var r=0;r<range.Rows;r++)
                    {
                        for (var c = 0; c < range.Columns; c++)
                        {
                            var v = range.Offset(r, c, 1, 1);
                            if (useText)
                            {
                                list.Add(v.Text);
                            }
                            else
                            {
                                double d=ConvertUtil.GetValueDouble(v.Value, false, true);
                                if (double.IsNaN((double)d)) d = 0;
                                if (isStacked)
                                {                                    
                                    var pv = GetSerieValue(prevList, i++);
                                    d += pv;
                                }
                                list.Add(d);
                            }
                        }
                    }
                }
            }
        }

        private double GetSerieValue(List<object> prevList, int i)
        {
            if (prevList.Count > i && ConvertUtil.IsExcelNumeric(prevList[i]))
            {
                var v=ConvertUtil.GetValueDouble(prevList[i], true, true);
                if (double.IsNaN(v)) return 0d;
                return v;
            }
            return 0d;
        }
    }
}
