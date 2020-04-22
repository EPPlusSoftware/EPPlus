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
using System;
using System.Xml;
using System.Globalization;
using OfficeOpenXml.Utils.Extentions;
namespace OfficeOpenXml.Drawing.Chart
{
    public sealed class ExcelChartAxisStandard : ExcelChartAxis
    {
        internal ExcelChartAxisStandard(ExcelChartBase chart, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string nsPrefix) : base(chart, nameSpaceManager, topNode, nsPrefix)
        {
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
                        return (eAxisTickMark)Enum.Parse(typeof(eAxisTickMark), v);
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
                        return (eAxisTickMark)Enum.Parse(typeof(eAxisTickMark), v);
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
        const string _formatPath = "c:numFmt/@formatCode";
        /// <summary>
        /// The Numberformat used
        /// </summary>
        public override string Format
        {
            get
            {
                return GetXmlNodeString(_formatPath);
            }
            set
            {
                SetXmlNodeString(_formatPath, value);
                if (string.IsNullOrEmpty(value))
                {
                    SourceLinked = true;
                }
                else
                {
                    SourceLinked = false;
                }
            }
        }
        const string _sourceLinkedPath = "c:numFmt/@sourceLinked";
        /// <summary>
        /// The Numberformats are linked to the source data.
        /// </summary>
        public override bool SourceLinked
        {
            get
            {
                return GetXmlNodeBool(_sourceLinkedPath);
            }
            set
            {
                SetXmlNodeBool(_sourceLinkedPath, value);
            }
        }
        ExcelChartTitle _title = null;
        /// <summary>
        /// Chart axis title
        /// </summary>
        public override ExcelChartTitle Title
        {
            get
            {
                if (_title == null)
                {
                    var node = TopNode.SelectSingleNode("c:title", NameSpaceManager);
                    if (node == null)
                    {
                        CreateNode("c:title");
                        node = TopNode.SelectSingleNode("c:title", NameSpaceManager);
                        node.InnerXml = ExcelChartTitle.GetInitXml("c");
                    }
                    _title = new ExcelChartTitle(_chart, NameSpaceManager, TopNode, "c");
                }
                return _title;
            }
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
                    DeleteNode(_minValuePath);
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
                    DeleteNode(_maxValuePath);
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
                    DeleteNode(_crossesAtPath);
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
                            DeleteNode(_custUnitPath);
                            SetXmlNodeString(_displayUnitPath, ((eBuildInUnits)value).ToString());
                            return;
                        }
                    }
                    DeleteNode(_displayUnitPath);
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
                    DeleteNode(_majorUnitPath);
                    DeleteNode(_majorUnitCatPath);
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
                    DeleteNode(_majorTimeUnitPath);
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
                    DeleteNode(_minorUnitPath);
                    DeleteNode(_minorUnitCatPath);
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
                    DeleteNode(_minorTimeUnitPath);
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
                    DeleteNode(_logbasePath);
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
    }
}
