/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// An axis for an extended chart
    /// </summary>
    public sealed class ExcelChartExAxis : ExcelChartAxis
    {
        internal ExcelChartExAxis(ExcelChart chart, XmlNamespaceManager nsm, XmlNode topNode) : base(chart, nsm, topNode, "cx")
        {
            SchemaNodeOrder = new string[] { "catScaling", "valScaling","title","units", "majorGridlines", "minorGridlines","majorTickMarks","minorTickMarks", "tickLabels", "numFmt", "spPr", "txPr" };
        }
        string _majorTickMarkPath = "cx:majorTickMarks/@type";
        /// <summary>
        /// Major tickmarks settings for the axis
        /// </summary>
        public override eAxisTickMark MajorTickMark 
        {
            get
            {
                return GetXmlNodeString(_majorTickMarkPath).ToEnum(eAxisTickMark.None);
            }
            set
            {
                SetXmlNodeString(_majorTickMarkPath, value.ToEnumString());
            }
        }
        string _minorTickMarkPath = "cx:majorTickMarks/@type";
        /// <summary>
        /// Minor tickmarks settings for the axis
        /// </summary>
        public override eAxisTickMark MinorTickMark
        {
            get
            {
                return GetXmlNodeString(_minorTickMarkPath).ToEnum(eAxisTickMark.None);
            }
            set
            {
                SetXmlNodeString(_minorTickMarkPath, value.ToEnumString());
            }
        }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override eAxisPosition AxisPosition 
        { 
            get
            {
                return eAxisPosition.Left;
            }
            internal set => throw new NotSupportedException(); 
        }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override eCrosses Crosses 
        { 
            get => eCrosses.AutoZero; 
            set => throw new NotSupportedException(); 
        }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override eCrossBetween CrossBetween 
        {
            get
            {
                return eCrossBetween.Between;
            } 
            set => throw new NotSupportedException(); 
        }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override double? CrossesAt 
        {
            get
            {
                return null;
            } 
            set => throw new NotSupportedException(); 
        }
        /// <summary>
        /// Labelposition. This property does not apply to extended charts.
        /// </summary>
        public override eTickLabelPosition LabelPosition 
        { 
            get => eTickLabelPosition.None; 
            set => throw new NotSupportedException(); 
        }
        /// <summary>
        /// If the axis is hidden. 
        /// </summary>
        public override bool Deleted 
        {
            get
            {
                return GetXmlNodeBool("@hidden");
            }
            set
            {
                SetXmlNodeBool("@hidden", value);
            }
        }
        /// <summary>
        /// Tick label position. This property does not apply to extended charts.
        /// </summary>
        public override eTickLabelPosition TickLabelPosition 
        {
            get
            {
                return eTickLabelPosition.None;
            }
            set => throw new NotSupportedException(); 
        }
        string _displayUnitPath = "cx:units/@unit";
        /// <summary>
        /// Display units. Please only use values in <see cref="eBuildInUnits"/> or 0 for none.
        /// </summary>
        public override double DisplayUnit 
        {
            get
            {
                var s=GetXmlNodeString(_displayUnitPath);
                if(string.IsNullOrEmpty(s))
                {
                    return 1;
                }
                try
                {
                    var e = Enum.Parse(typeof(eBuildInUnits), s);
                    return (double)e;
                }
                catch
                {
                    return 0;
                }
            }
            set
            {
                if(value==0 || value==1)
                {
                    DeleteNode("cx:units");
                }
                try
                {
                    var e = (eBuildInUnits)value;
                    SetXmlNodeString("", e.ToEnumString());
                }
                catch
                {
                    throw new InvalidOperationException("DisplayUnit property for extended charts can only contain Build in Units, matching the eBuildInUnits enum or be 0 for no units");
                }
            }
        }
        /// <summary>
        /// The title of the chart
        /// </summary>
        public new ExcelChartExTitle Title
        {
            get
            {
                return (ExcelChartExTitle)GetTitle();
            }
        }
        internal override ExcelChartTitle GetTitle()
        {
            if (_title == null)
            {
                var node = AddTitleNode();
                _title = new ExcelChartExTitle(_chart, NameSpaceManager, node);
            }
            return _title;
        }

        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override double? MinValue 
        {
            get
            {
                return null;
            }
            set => throw new NotSupportedException(); 
        }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override double? MaxValue { get => null; set => throw new NotSupportedException(); }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override double? MajorUnit { get => null; set => throw new NotSupportedException(); }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override eTimeUnit? MajorTimeUnit { get => null; set => throw new NotSupportedException(); }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override double? MinorUnit 
        {
            get
            {
                return null;
            }
            set => throw new NotSupportedException(); }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override eTimeUnit? MinorTimeUnit 
        {
            get
            {
                return null;
            }
            set => throw new NotSupportedException(); }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override double? LogBase 
        {
            get
            {
                return null;
            }
            set => throw new NotSupportedException(); 
        }
        /// <summary>
        /// This property is not used for extended charts. Trying to set this property will result in a NotSupportedException.
        /// </summary>
        public override eAxisOrientation Orientation 
        {
            get
            {
                return eAxisOrientation.MinMax;
            }
            set => throw new NotSupportedException(); 
        }
        
        internal override string Id
        {
            get
            {
                return GetXmlNodeString("@id");
            }
        }
        internal override eAxisType AxisType
        {
            get
            {
                if(Id=="0")
                {
                    return eAxisType.Cat;
                }
                else
                {
                    return eAxisType.Val;
                }
            }
        }
    }
}
    
