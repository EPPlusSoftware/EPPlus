using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public sealed class ExcelChartExAxis : ExcelChartAxis
    {
        public ExcelChartExAxis(ExcelChart chart, XmlNamespaceManager nsm, XmlNode topNode) : base(chart, nsm, topNode, "cx")
        {

        }
        string _majorTickMarkPath = "cx:majorTickMarks/@type";
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
        /// Where the axis is located. This property does not apply to extended charts
        /// </summary>
        public override eAxisPosition AxisPosition 
        { 
            get
            {
                return eAxisPosition.Left;
            }
            internal set => throw new NotImplementedException(); 
        }
        public override eCrosses Crosses 
        { 
            get => eCrosses.AutoZero; 
            set => throw new NotImplementedException(); 
        }
        public override eCrossBetween CrossBetween 
        {
            get
            {
                return eCrossBetween.Between;
            } 
            set => throw new NotImplementedException(); 
        }
        public override double? CrossesAt 
        {
            get
            {
                return null;
            } 
            set => throw new NotImplementedException(); 
        }
        public override eTickLabelPosition LabelPosition 
        { 
            get => eTickLabelPosition.None; 
            set => throw new NotImplementedException(); 
        }
        public override bool Deleted 
        { 
            get => false; 
            set => throw new NotImplementedException(); 
        }
        public override eTickLabelPosition TickLabelPosition 
        {
            get
            {
                return eTickLabelPosition.None;
            }
            set => throw new NotImplementedException(); 
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

        public new ExcelChartExTitle Title
        {
            get
            {
                if(_title==null)
                {
                    AddTitleNode();
                    _title = new ExcelChartExTitle(_chart, NameSpaceManager, TopNode);
                }
                return (ExcelChartExTitle)_title;
            }
        }

        public override double? MinValue 
        {
            get
            {
                return null;
            }
            set => throw new NotImplementedException(); 
        }
        public override double? MaxValue { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override double? MajorUnit { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override eTimeUnit? MajorTimeUnit { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override double? MinorUnit 
        {
            get
            {
                return null;
            }
            set => throw new NotImplementedException(); }
        public override eTimeUnit? MinorTimeUnit 
        {
            get
            {
                return null;
            }
            set => throw new NotImplementedException(); }
        public override double? LogBase 
        {
            get
            {
                return null;
            }
            set => throw new NotImplementedException(); }
        public override eAxisOrientation Orientation 
        {
            get
            {
                return eAxisOrientation.MinMax;
            }
            set => throw new NotImplementedException(); 
        }

        internal override string Id
        {
            get
            {
                return GetXmlNodeString("@id");
            }
        }
    }
}
