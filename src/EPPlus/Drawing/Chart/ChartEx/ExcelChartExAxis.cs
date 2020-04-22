using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public sealed class ExcelChartExAxis : ExcelChartAxis
    {
        public ExcelChartExAxis(ExcelChartBase chart, XmlNamespaceManager nsm, XmlNode topNode) : base(chart, nsm, topNode, "cx")
        {

        }

        public override eAxisTickMark MajorTickMark { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override eAxisTickMark MinorTickMark { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override eAxisPosition AxisPosition { get => throw new NotImplementedException(); internal set => throw new NotImplementedException(); }
        public override eCrosses Crosses { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override eCrossBetween CrossBetween { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override double? CrossesAt { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override string Format { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override bool SourceLinked { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override eTickLabelPosition LabelPosition { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override bool Deleted { get => false; set => throw new NotImplementedException(); }
        public override eTickLabelPosition TickLabelPosition { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override double DisplayUnit { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public override ExcelChartTitle Title => throw new NotImplementedException();

        public override double? MinValue { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override double? MaxValue { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override double? MajorUnit { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override eTimeUnit? MajorTimeUnit { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override double? MinorUnit { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override eTimeUnit? MinorTimeUnit { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override double? LogBase { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public override eAxisOrientation Orientation { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        internal override string Id => throw new NotImplementedException();
    }
}
