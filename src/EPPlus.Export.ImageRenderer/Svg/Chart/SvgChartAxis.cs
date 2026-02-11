/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Svg.Chart.Util;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Text;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartAxis : SvgChartObject, IDrawingChartAxis
    {
        internal SvgChartAxis(SvgChart sc, ExcelChartAxisStandard ax) : base(sc)
        {
            SvgChart= sc;
            Axis = ax;
            SetMargins(ax.TextBody);

            if (sc.Chart.Series.Count == 0 || (ax.Deleted == true && ax.Title == null))
            {
                return;
            }

            if(ax.HasTitle)
            {
                Title = new SvgChartTitle(sc, ax.Title, "Axis Title", this);
            }
            else
            {
                Title = null;
            }

            if (ax.Deleted == false || ax.HasMajorGridlines || ax.HasMinorGridlines)
            {
                Values = GetAxisValue(ax, Rectangle, out double? min, out double? max, out double? majorUnit);
                AxisValues = GetAxisDisplayValues(ax, Values, min, max, majorUnit);
                
                Min = min ?? 0D;
                Max = max ?? (Values.Count > 0 ? ConvertUtil.GetValueDouble(Values[Values.Count - 1], false, true) : 0D);
                MajorUnit = majorUnit ?? 1;
                MinorUnit = ax.MinorUnit ?? GetAutoMinUnit(MajorUnit);
                if (ax.Deleted == false)
                {
                    if (ax.Layout.HasLayout)
                    {
                        Rectangle = GetRectFromManualLayout(sc, ax.Layout);
                    }
                    else
                    {
                        Rectangle = new SvgRenderRectItem(sc, sc.Bounds);
                        if (ax.AxisPosition == eAxisPosition.Left || ax.AxisPosition == eAxisPosition.Right)
                        {
                            if (ax.AxisPosition == eAxisPosition.Left)
                            {
                                Rectangle.Width = GetTextWidest(sc, ax) + LeftMargin;
                                var ll = 8D;
                                if (sc.Chart.Legend.Position == eLegendPosition.Left)
                                {
                                    ll = sc.Legend.Rectangle.Right + sc.Legend.RightMargin;
                                }
                                Rectangle.Left = Title == null ? ll : Title.Rectangle.Right;
                            }
                            else
                            {
                                Rectangle.Width = GetTextWidest(sc, ax) + RightMargin;
                                var lp = sc.ChartArea.Rectangle.Width - Rectangle.Width - 8D;
                                if (sc.Chart.Legend.Position == eLegendPosition.Right)
                                {
                                    lp = sc.Legend.Rectangle.Left + -Rectangle.Width;
                                }
                                Rectangle.Left = Title == null ? lp : Title.Rectangle.Left - Rectangle.Width;
                            }
                        }
                        else
                        {
                            Rectangle.Height = GetTextHeight(sc, ax);
                            Rectangle.Top = Title == null || ax.AxisPosition == eAxisPosition.Top ? sc.ChartArea.Rectangle.Height - 8 - Rectangle.Height : Title.Rectangle.Top - Rectangle.Height - 8;
                        }
                    }

                    Rectangle.FillColor = "none";

                    Line = new SvgRenderLineItem(sc, Rectangle.Bounds);
                    Line.SetDrawingPropertiesBorder(ax.Border, sc.Chart.StyleManager.Style.Title.BorderReference.Color, ax.Border.Fill.Style != eFillStyle.NoFill, 0.75);
                }
            }
        }

        private double GetAutoMinUnit(double majorUnit)
        {
            return majorUnit / 5;
        }

        internal ExcelChartAxis Axis { get; }
        internal SvgChart SvgChart { get; }
        private List<string> GetAxisDisplayValues(ExcelChartAxisStandard ax, List<object> values, double? min, double? max, double? majorUnit)
        {
            var displayValues = new List<string>();
            var nf = new ExcelFormatTranslator(ax.Format, 0);
            //Excel replaces the format with a default date format if the axis is date based.
            if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
            {
                if(ax.Format == "m/d/yyyy")
                {
                    var sdFormat = ExcelNumberFormat.GetFromBuildInFromID(14); //14 is standard regional short date.
                    nf = new ExcelFormatTranslator(sdFormat, 14);
                }
            }
            foreach (var v in values)
            {
                var s = ValueToTextHandler.FormatValue(v, false, nf, null, out bool isValidFormat);
                displayValues.Add(s);
            }
            return displayValues;
        }
        private double GetTextHeight(SvgChart sc, ExcelChartAxisStandard ax)
        {
            var tm = sc.Chart.WorkSheet._package.Settings.TextSettings.GenericTextMeasurerTrueType;
            var highest = 0f;
            var mf = ax.Font.GetMeasureFont();
            foreach (var s in AxisValues)
            {
                var m = tm.MeasureText(s, mf);
                if (m.Height > highest)
                {
                    highest = m.Height;
                }
            }
            return highest.PointToPixel();
        }

        private double GetTextWidest(SvgChart sc, ExcelChartAxisStandard ax)
        {
            var tm = sc.Chart.WorkSheet._package.Settings.TextSettings.GenericTextMeasurerTrueType;
            
            var widest = 0f;
            var mf = ax.Font.GetMeasureFont();
            foreach(var s in AxisValues)
            {
                var m= tm.MeasureText(s, mf);
                if (m.Width > widest)
                {
                    widest = m.Width;
                }
            }
            return widest.PointToPixel();
        }

        public List<object> Values
        {
            get;
            private set;
        }
        public List<string> AxisValues { get; private set; }

        public List<SvgRenderLineItem> MajorAxisPositions { get; private set; }
        public List<SvgRenderLineItem> MinorAxisPositions { get; private set; }
        public List<SvgRenderLineItem> MajorGridlinePositions { get; private set; }
        public List<SvgRenderLineItem> MinorGridlinePositions { get; private set; }
        public List<SvgTextBox> AxisValuesTextBoxes
        {
            get;
            private set;
        }

        public SvgChartTitle Title { get; set; }
        public double Min { get; set; }
        public double Max { get; set; }
        public double MajorUnit { get; set; }
        public double MinorUnit { get; set; }
        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            Title?.AppendRenderItems(renderItems);
            //Title?.Render(sb);
            if(Rectangle!=null) renderItems.Add(Rectangle);

            if (MajorGridlinePositions != null)
            {
                foreach (var tm in MajorGridlinePositions)
                {
                    renderItems.Add(tm);
                }
            }
            if (MinorGridlinePositions != null)
            {
                foreach (var tm in MinorAxisPositions)
                {
                    renderItems.Add(tm);
                }
            }

            if (Line != null) renderItems.Add(Line);

            if (MajorAxisPositions != null)
            {
                foreach (var tm in MajorAxisPositions)
                {
                    renderItems.Add(tm);
                }
            }
            if (MinorAxisPositions != null)
            {
                foreach (var tm in MinorAxisPositions)
                {
                    renderItems.Add(tm);
                }
            }

            if (AxisValuesTextBoxes != null && AxisValuesTextBoxes.Count > 0)
            {
                foreach (var tb in AxisValuesTextBoxes)
                {
                    tb.AppendRenderItems(renderItems);
                }
            }
            
        }


        internal void AddTickmarksAndValues()
        {            
            if (Axis.MajorTickMark != eAxisTickMark.None)
            {
                MajorAxisPositions = AddTickmarks(MajorUnit,double.NaN, 4, Axis.MajorTickMark);
            }

            if (Axis.MinorTickMark != eAxisTickMark.None)
            {
                MinorAxisPositions = AddTickmarks(MinorUnit, MajorUnit, 2, Axis.MinorTickMark);
            }

            if(Axis.HasMajorGridlines)
            {
                MajorGridlinePositions = AddGridlines(MajorUnit, double.NaN, Axis.MajorGridlines, Chart.StyleManager.Style.GridlineMajor);
            }

            if ((Axis.HasMinorGridlines))
            {
                MinorGridlinePositions = AddGridlines(MinorUnit, MajorUnit, Axis.MinorGridlines, Chart.StyleManager.Style.GridlineMinor);
            }

            if (Axis.CrossBetween == eCrossBetween.MidCat)
            {
                //TODO: Adjust for Crossbetween 
            }
            else
            {

            }
            if (AxisValues != null && AxisValues.Count > 0 && Axis.Deleted==false)
            {
                AxisValuesTextBoxes = GetAxisValueTextBoxes();
            }
        }

        private List<SvgTextBox> GetAxisValueTextBoxes()
        {
            var tm = Chart.WorkSheet._package.Settings.TextSettings.GenericTextMeasurerTrueType;
            var mf = Axis.Font.GetMeasureFont();
            var axisStyle = GetAxisStyleEntry();
            var ret= new List<SvgTextBox>();
            double maxWidth, maxHeight;
            if(Axis.AxisPosition==eAxisPosition.Left || Axis.AxisPosition == eAxisPosition.Right)
            {
                maxWidth = SvgChart.ChartArea.Rectangle.Width / 3; //TODO: Check this value.
                maxHeight = Rectangle.Height / AxisValues.Count;
            }
            else
            {
                maxWidth = Rectangle.Width / AxisValues.Count;
                maxHeight = SvgChart.ChartArea.Bounds.Height / 3; //TODO: Check this value.
            }
            for (var i = 0; i < AxisValues.Count; i++)
            {
                var v = AxisValues[i];
                var m = tm.MeasureText(v, mf);
                var x = GetAxisItemLeft(i, m);
                var y = GetAxisItemTop(i, m);
                //var tb = new TextBox(Chart, x, y, m.Width, m.Height);
                //var bounds = new BoundingBox();
                //bounds.Left = x;
                //bounds.Top = y;
                var width = m.Width.PointToPixel();
                var height = m.Height.PointToPixel();
                var tb = new SvgTextBox(SvgChart, Rectangle.Bounds, x, y, width, height, maxWidth, maxHeight);

                var p = Axis.TextBody.Paragraphs.FirstOrDefault();
                tb.TextBody.ImportParagraph(p, 0, v);

                //tb.TextBody.Paragraphs[0].AddText(v, Axis.Font);
                tb.Rectangle.SetDrawingPropertiesFill(Axis.Fill, axisStyle.FillReference.Color);
                ret.Add(tb);
            }
            return ret;
        }

        private double GetAxisItemLeft(int i, OfficeOpenXml.Interfaces.Drawing.Text.TextMeasurement m)
        {
            if (Axis.AxisPosition == eAxisPosition.Left)
                {
                return Rectangle.Left;
            }
            else if (Axis.AxisPosition == eAxisPosition.Right)
            {

                return Rectangle.Left;
            }
            else
            {
                if (Axis.AxisType == eAxisType.Cat)
                {
                    var majorWidth = Rectangle.Width / AxisValues.Count;
                    return Rectangle.Left + majorWidth * i + (majorWidth / 2) - m.Width.PointToPixel() / 2;
                }
                else
                {
                    var majorWidth = Rectangle.Width / (AxisValues.Count - 1);

                    return Rectangle.Left + majorWidth * i - m.Width.PointToPixel() / 2;
                }
            }
        }

        private double GetAxisItemTop(int i, OfficeOpenXml.Interfaces.Drawing.Text.TextMeasurement m)
        {
            if (Axis.AxisPosition == eAxisPosition.Top)
            {
                return Rectangle.Top - m.Height.PointToPixel() - TopMargin;
            }
            else if (Axis.AxisPosition == eAxisPosition.Bottom)
            {
                return Rectangle.Bottom - m.Height.PointToPixel() + BottomMargin;
            }
            else
            {
                var majorHeight = Rectangle.Height / (AxisValues.Count-1);
                if (Axis.AxisType == eAxisType.Cat)
                {
                    return Rectangle.Top + majorHeight * (AxisValues.Count - i - 1) + (majorHeight / 2) - m.Height.PointToPixel() / 2;
                }
                else
                {
                    return Rectangle.Top + majorHeight * (AxisValues.Count - i - 1) - m.Height.PointToPixel() / 2;
                }
            }

        }

        private List<SvgRenderLineItem> AddTickmarks(double units, double parentUnit, float tickMarkWidth, eAxisTickMark type)
        {
            var axisStyle = GetAxisStyleEntry();

            var tms = new List<SvgRenderLineItem>();
            double min;
            if (Axis.AxisType == eAxisType.Cat)
            {
                min = 0;
            }
            else
            {
                min = Min;
            }
            float tickMarkWidthInside=0, tickMarkWidthOutside=0;
            if(type==eAxisTickMark.In || type==eAxisTickMark.Cross)
            {
                tickMarkWidthInside = tickMarkWidth;
            }
            if(type==eAxisTickMark.Out|| type == eAxisTickMark.Cross)
            {
                tickMarkWidthOutside = tickMarkWidth;
            }

            var diff = Max - min;
            for (double d = min; d <= Max; d += units)
            {
                if (double.IsNaN(parentUnit) || (d % parentUnit != 0))
                {
                    float x1, y1, x2, y2;
                    switch (Axis.AxisPosition)
                    {
                        case eAxisPosition.Left:
                            y1 = (float)(Rectangle.Top + Rectangle.Height - ((d - min) / diff * Rectangle.Height));
                            y2 = y1;                            
                            x1 = (float)Rectangle.Right - tickMarkWidthInside;
                            x2 = (float)Rectangle.Right + tickMarkWidthOutside;
                            break;
                        case eAxisPosition.Right:
                            y1 = (float)(Rectangle.Top + Rectangle.Height - ((d - min) / diff * Rectangle.Height));
                            y2 = y1;
                            x1 = (float)Rectangle.Left - tickMarkWidthInside;
                            x2 = (float)Rectangle.Left + tickMarkWidthOutside;
                            break;
                        case eAxisPosition.Top:
                            x1 = (float)(Rectangle.Left + ((d - min) / diff * Rectangle.Width));
                            x2 = x1;
                            y1 = (float)Rectangle.Bottom - tickMarkWidthInside;
                            y2 = (float)Rectangle.Bottom + tickMarkWidthOutside;
                            break;
                        case eAxisPosition.Bottom:
                            x1 = (float)(Rectangle.Left + ((d - min) / diff * Rectangle.Width));
                            x2 = x1;
                            y1 = (float)Rectangle.Top - tickMarkWidthInside;
                            y2 = (float)Rectangle.Top + tickMarkWidthOutside;
                            break;
                        default:
                            throw new InvalidOperationException("Invalid axis position");
                    }
                    var tm = new SvgRenderLineItem(SvgChart, SvgChart.Bounds);
                    tm.X1 = x1;
                    tm.Y1 = y1;
                    tm.X2 = x2;
                    tm.Y2 = y2;
                    tm.SetDrawingPropertiesBorder(Axis.Border, axisStyle.BorderReference.Color, true);
                    tms.Add(tm);
                }
            }
            return tms;
        }
        private List<SvgRenderLineItem> AddGridlines(double units, double parentUnit, ExcelDrawingBorder lineItem, ExcelChartStyleEntry styleEntry)
        {
            var axisStyle = GetAxisStyleEntry();

            var tms = new List<SvgRenderLineItem>();
            double min;
            if (Axis.AxisType == eAxisType.Cat)
            {
                min = 0;
            }
            else
            {
                min = Min;
            }
            var pa = SvgChart.Plotarea;
            var diff = Max - min;
            for (double d = min; d <= Max; d += units)
            {
                if(d==min && Line!=null && Line.BorderWidth>0) continue;
                if (double.IsNaN(parentUnit) || (d % parentUnit != 0))
                {
                    float x1, y1, x2, y2;
                    switch (Axis.AxisPosition)
                    {
                        case eAxisPosition.Left:
                        case eAxisPosition.Right:
                            y1 = (float)(pa.Rectangle.Top + pa.Rectangle.Height - ((d - min) / diff * pa.Rectangle.Height));
                            y2 = y1;
                            x1 = (float)pa.Rectangle.Right;
                            x2 = (float)pa.Rectangle.Left;
                            break;
                        case eAxisPosition.Top:
                        case eAxisPosition.Bottom:
                            x1 = (float)(pa.Rectangle.Left + ((d - min) / diff * pa.Rectangle.Width));
                            x2 = x1;
                            y1 = (float)pa.Rectangle.Top;
                            y2 = (float)pa.Rectangle.Bottom;
                            break;
                        default:
                            throw new InvalidOperationException("Invalid axis position");
                    }

                    var tm = new SvgRenderLineItem(SvgChart, SvgChart.Bounds);
                    tm.X1 = x1;
                    tm.Y1 = y1;
                    tm.X2 = x2;
                    tm.Y2 = y2;
                    tm.SetDrawingPropertiesBorder(lineItem, styleEntry.BorderReference.Color, true, lineItem.Width);
                    tms.Add(tm);
                }
            }
            return tms;
        }

        private ExcelChartStyleEntry GetAxisStyleEntry()
        {
            ExcelChartStyleEntry axisStyle;
            switch (Axis.AxisType)
            {
                case eAxisType.Cat:
                    axisStyle = Chart.StyleManager.Style.CategoryAxis;
                    break;
                case eAxisType.Serie:
                    axisStyle = Chart.StyleManager.Style.SeriesAxis;
                    break;
                default:
                    axisStyle = Chart.StyleManager.Style.ValueAxis;
                    break;
            }

            return axisStyle;
        }

        internal double GetPositionInPlotarea(double val)
        {
            if (Axis.AxisPosition == eAxisPosition.Left || Axis.AxisPosition == eAxisPosition.Right)
            {
                if (Axis.AxisType == eAxisType.Cat)
                {
                    var majorHeight = SvgChart.Plotarea.Rectangle.Height / Max;
                    return (majorHeight * val + (majorHeight / 2));
                }
                else
                {
                    if (val < Min || val > Max) return double.NaN;
                    var diff = Max - Min;
                    return (((Max-val) / diff * SvgChart.Plotarea.Rectangle.Height));
                }
            }
            else
            {
                if (Axis.AxisType == eAxisType.Cat)
                {
                    var majorWidth = SvgChart.Plotarea.Rectangle.Width / Max;
                    return (majorWidth * val + (majorWidth / 2));
                }
                else
                {
                    if (val < Min || val > Max) return double.NaN;
                    var diff = Max - Min;
                    return (((val-Min) / diff * SvgChart.Plotarea.Rectangle.Width));
                }
            }
        }
        protected List<object> GetAxisValue(ExcelChartAxisStandard ax, RenderItem rect, out double? min, out double? max, out double? majorUnit)
        {
            var values = ax.GetAxisValues(out bool isCount);
            if (ax.AxisType == eAxisType.Cat &&
                isCount == false)
            {
                min = 0;
                max = values.Length;
                majorUnit = 1;
                return values.ToList();
            }
            var l = new List<object>();
            min = double.MaxValue;
            max = double.MinValue;
            foreach (var v in values)
            {
                var d = ConvertUtil.GetValueDouble(v, false, true);
                if (double.IsNaN(d))
                {
                    d = 0;
                }
                if (min > d)
                {
                    min = d;
                }
                if (max < d)
                {
                    max = d;
                }
            }
            var options = new AxisOptions
            {
                LockedMin = ax.MinValue,
                LockedMax = ax.MaxValue,
                LockedInterval = ax.MajorUnit,
                LockedIntervalUnit = ax.MajorTimeUnit,
                AddPadding = ax.AxisPosition == eAxisPosition.Left || ax.AxisPosition == eAxisPosition.Right,
                Axis = ax,
                IsStacked100 = Chart.IsTypePercentStacked()
            };

            var length = ax.AxisPosition == eAxisPosition.Left || ax.AxisPosition == eAxisPosition.Right ? SvgChart.Bounds.Height : SvgChart.Bounds.Width; //Fix and use plotarea width/height.
            if(isCount)
            {
                majorUnit = 1;
                for(int i=1;i<=max;i++)
                {
                    l.Add(i);
                }
                return l;
            }
            if (ax.IsDate)
            {
                var res = DateAxisScaleCalculator.Calculate(min ?? 0, max ?? 0, length, options);
                var dt = DateTime.FromOADate(res.Min);
                var maxDt = DateTime.FromOADate(res.Max);
                while (dt < maxDt)
                {
                    l.Add(dt);
                    switch(res.MajorDateUnit ?? eTimeUnit.Days)
                    {
                        case eTimeUnit.Years:
                            dt = dt.AddYears((int)res.MajorInterval);
                            break;
                        case eTimeUnit.Months:
                            dt = dt.AddMonths((int)res.MajorInterval);
                            break;
                        case eTimeUnit.Days:
                            dt = dt.AddDays((int)res.MajorInterval);
                            break;
                    }
                }

                min = res.Min;
                max = res.Max;
                majorUnit = res.MajorInterval;
            }
            else
            {
                var res = ValueAxisScaleCalculator.Calculate(min ?? 0, max ?? 0, length, options);
                for (var v = res.Min; v <= res.Max; v += res.MajorInterval)
                {
                    l.Add(v);
                }

                min = res.Min;
                max = res.Max;
                majorUnit = res.MajorInterval;
            }

            return l;
        }
    }
}