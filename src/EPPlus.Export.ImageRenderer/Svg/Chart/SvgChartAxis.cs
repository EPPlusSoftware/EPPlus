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
using EPPlus.Export.ImageRenderer;
using EPPlus.Export.ImageRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Svg.Chart.Util;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartAxis : SvgChartObject, IDrawingChartAxis
    {
        private const double COS45 = 0.70710678118654757; //Constant for Math.Sin(Math.PI / 4) --45 degrees
        internal SvgChartAxis(SvgChart sc, ExcelChartAxisStandard ax) : base(sc)
        {
            SvgChart = sc;
            Axis = ax;
            SetMargins(ax.TextBody);

            if (sc.Chart.Series.Count == 0 || (ax.Deleted == true && ax.HasTitle==false))
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
                Values = GetAxisValue(ax, sc.ChartArea.Rectangle, out double? min, out double? max, out double? majorUnit, out eTimeUnit? dateUnit, out eTextOrientation orientation);
                AxisValues = GetAxisDisplayValues(ax, Values, min, max, majorUnit);
                
                Min = min ?? 0D;
                Max = max ?? (Values.Count > 0 ? ConvertUtil.GetValueDouble(Values[Values.Count - 1], false, true) : 0D);
                MajorUnit = majorUnit ?? 1;
                MinorUnit = ax.MinorUnit ?? GetAutoMinUnit(MajorUnit);
                MajorDateUnit = dateUnit;
                LabelOrientation = orientation;
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
                                if (sc.Chart.HasLegend  && sc.Chart.Legend.Position == eLegendPosition.Left)
                                {
                                    ll = sc.Legend.Rectangle.Right + sc.Legend.RightMargin;
                                }
                                Rectangle.Left = Title == null ? ll : Title.Rectangle.Right;
                            }
                            else
                            {
                                Rectangle.Width = GetTextWidest(sc, ax) + RightMargin;
                                var lp = sc.ChartArea.Rectangle.Width - Rectangle.Width - 8D;
                                if (sc.Chart.HasLegend && sc.Chart.Legend.Position == eLegendPosition.Right)
                                {
                                    lp = sc.Legend.Rectangle.Left + -Rectangle.Width;
                                }
                                Rectangle.Left = Title == null ? lp : Title.Rectangle.Left - Rectangle.Width-LeftMargin;
                            }
                        }
                        else
                        {
                            Rectangle.Height = GetTextHeight(sc, ax);
                            //TODO:Fix
                            Rectangle.Top = Title == null || ax.AxisPosition == eAxisPosition.Top ? sc.ChartArea.Rectangle.Height - 8 - Rectangle.Height : Title.Rectangle.Top - Rectangle.Height - 8;
                        }
                    }

                    Rectangle.FillColor = "none";

                    Line = new SvgRenderLineItem(sc, Rectangle.Bounds);
                    Line.SetDrawingPropertiesBorder(ax.Border, sc.Chart.StyleManager.Style.Title.BorderReference.Color, ax.Border.Fill.Style != eFillStyle.NoFill, 1);
                    if(Line.BorderWidth < 1)
                    {
                        Line.BorderWidth = 1;
                    }
                }
            }
        }

        private double GetAutoMinUnit(double majorUnit)
        {
            return majorUnit / 5;
        }

        internal ExcelChartAxis Axis { get; }
        internal SvgChart SvgChart { get; }
        internal SvgRenderLineItem Line { get; set; }
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
            foreach (var v in CategoryAxisScaleCalculator.GetUniqueValues(values))
            {
                var s = ValueToTextHandler.FormatValue(v, false, nf, null, out bool isValidFormat);
                displayValues.Add(s);
            }
            return displayValues;
        }
        private double GetTextHeight(SvgChart sc, ExcelChartAxisStandard ax)
        {
            var tm = sc.TextMeasurer;
            var highest = 0D;
            var mf = ax.Font.GetMeasureFont();
            foreach (var s in AxisValues)
            {
                var m = tm.MeasureText(s, mf);
                switch (LabelOrientation)
                {
                    case eTextOrientation.Horizontal:
                        if (m.Height > highest)
                        {
                            highest = m.Height;
                        }
                        break;
                    case eTextOrientation.Diagonal:
                        var width = (m.Width + m.Height) * COS45;
                        if (width > highest)
                        {
                            highest = width;
                        }
                        break;
                    case eTextOrientation.Vertical:
                        if (m.Width > highest)
                        {
                            highest = m.Width;
                        }
                        break;
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
        public List<RenderItem> MajorGridlinePositions { get; private set; }
        public List<RenderItem> MinorGridlinePositions { get; private set; }
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
        public eTimeUnit? MajorDateUnit { get; set; }
        public eTextOrientation LabelOrientation { get; set; }
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


        internal void AddTickmarksAndValues(List<RenderItem> DefItems)
        {
            if (Axis.MajorTickMark != eAxisTickMark.None)
            {
                MajorAxisPositions = AddTickmarks(MajorUnit, MajorDateUnit, double.NaN, 4D.PixelToPoint(), Axis.MajorTickMark);
            }

            if (Axis.MinorTickMark != eAxisTickMark.None)
            {
                MinorAxisPositions = AddTickmarks(MinorUnit, MajorDateUnit, MajorUnit, 2D.PixelToPoint(), Axis.MinorTickMark);
            }

            if(Axis.HasMajorGridlines)
            {
                MajorGridlinePositions = AddGridlines(MajorUnit, double.NaN, Axis.MajorGridlines, Chart.StyleManager.Style.GridlineMajor);
                //DefItems.Add(MajorGridlinePositions[0]);
                //MajorGridlinePositions.RemoveAt(0);
            }

            if ((Axis.HasMinorGridlines))
            {
                MinorGridlinePositions = AddGridlines(MinorUnit, MajorUnit, Axis.MinorGridlines, Chart.StyleManager.Style.GridlineMinor);
                //DefItems.Add(MinorGridlinePositions[0]);
                //MinorGridlinePositions.RemoveAt(0);
            }

            if (Axis.CrossBetween == eCrossBetween.MidCat)
            {
                //TODO: Adjust for Crossbetween 
            }
            else
            {

            }
            if (AxisValues != null && AxisValues.Count > 0 && Axis.Deleted==false && Axis.LabelPosition != eTickLabelPosition.None)
            {
                AxisValuesTextBoxes = GetAxisValueTextBoxes();
            }
        }

        private List<SvgTextBox> GetAxisValueTextBoxes()
        {
            var ret = new List<SvgTextBox>();
            if (Axis.LabelPosition == eTickLabelPosition.None) return ret;

            var tm = Chart.WorkSheet._package.Settings.TextSettings.GenericTextMeasurerTrueType;
            var mf = Axis.Font.GetMeasureFont();
            var axisStyle = GetAxisStyleEntry();
            double maxWidth, maxHeight;
            if(Axis.AxisPosition==eAxisPosition.Left || Axis.AxisPosition == eAxisPosition.Right)
            {
                maxWidth = SvgChart.ChartArea.Rectangle.Width / 3; //TODO: Check this value.
                maxHeight = Rectangle.Height / AxisValues.Count;
            }
            else
            {
                switch (LabelOrientation)
                {
                    case eTextOrientation.Vertical:
                        maxWidth =  SvgChart.ChartArea.Rectangle.Height / 3;
                        maxHeight = Rectangle.Width / AxisValues.Count; //TODO: Check this value.
                        break;                    
                    case eTextOrientation.Diagonal:
                        maxWidth = (Rectangle.Width + Rectangle.Height) / COS45;
                        maxHeight = SvgChart.ChartArea.Rectangle.Height / 3; //TODO: Check this value.
                        break;
                    default:
                        maxWidth = Rectangle.Width / AxisValues.Count;
                        maxHeight = SvgChart.ChartArea.Rectangle.Height / 3; //TODO: Check this value.
                        break;
                }
            }
            double widest=0;
            for (var i = 0; i < AxisValues.Count; i++)
            {

                var v = AxisValues[i];
                var m = tm.MeasureText(v, mf);
                var ticMarkX = GetAxisItemLeft(i, m);
                var ticMarkY = GetAxisItemTop(i, m);
                var width = m.Width;
                var height = m.Height;
                double x, y;
                if(LabelOrientation==eTextOrientation.Horizontal)
                {
                    if (Axis.AxisType == eAxisType.Cat)
                    {
                        x = ticMarkX;
                        y = ticMarkY;
                    }
                    else
                    {
                        if(Axis.IsVertical)
                        {
                            x = ticMarkX;
                            y = ticMarkY - height / 2;
                        }
                        else
                        {
                            x = ticMarkX; // - width / 2;
                            y = ticMarkY;
                        }
                    }
                }
                else
                {
                    var rot = LabelOrientation == eTextOrientation.Diagonal ? -45 : -90;
                    double cos = Math.Cos(MathHelper.Radians(rot));
                    double sin = Math.Sin(MathHelper.Radians(rot));

                    if (LabelOrientation == eTextOrientation.Diagonal)
                    {
                        x = ticMarkX - (height / 2) * cos;
                        if (Axis.AxisPosition == eAxisPosition.Bottom)
                        {
                            y = ticMarkY + 4 + TopMargin - (height / 2 * cos);
                        }
                        else //Top
                        {
                            y = ticMarkY - 4 - BottomMargin - (height/2 * cos);
                        }
                    }
                    else
                    {
                        x = ticMarkX - (height / 2);
                        if (Axis.AxisPosition == eAxisPosition.Bottom)
                        {
                            y = ticMarkY + BottomMargin + 4;
                        }
                        else //Top
                        {
                            y = ticMarkY - TopMargin - 4;
                        }
                    }
                }
                
                var tb = new SvgTextBox(SvgChart, Rectangle.Bounds, x, y, width, height, maxWidth, maxHeight);
                if (LabelOrientation == eTextOrientation.Diagonal)
                {
                    tb.Rotation = -45;
                    if(Axis.AxisPosition==eAxisPosition.Bottom)
                    {
                        tb.TextAnchor = eTextAnchor.End;
                    }
                }

                else if (LabelOrientation == eTextOrientation.Vertical)
                {
                    tb.Rotation = -90;
                    if (Axis.AxisPosition == eAxisPosition.Bottom)
                    {
                        tb.TextAnchor = eTextAnchor.End;
                    }
                }

                var p = Axis.TextBody.Paragraphs.FirstOrDefault();

                if (p.HorizontalAlignment != eTextAlignment.Center && (Axis.AxisPosition == eAxisPosition.Bottom || Axis.AxisPosition == eAxisPosition.Top))
                {
                    //Horizontal axises are always center aligned visually
                    //Should be broken out as input to ImportParagraph instead of changing the base item
                    p.HorizontalAlignment = eTextAlignment.Center;
                }

                tb.TextBody.ImportParagraph(p, 0, v);

                //tb.TextBody.Paragraphs[0].AddText(v, Axis.Font);
                tb.Rectangle.SetDrawingPropertiesFill(Axis.Fill, axisStyle.FillReference.Color);

                if(widest < tb.Width)
                {
                    widest = tb.Width;
                }
                ret.Add(tb);
            }

            if(Axis.IsVertical)
            {
                //If the axis is vertical, we need to adjust the left position of the textboxes to align them to the right and not have them overlap with the axis line.
                if (Axis.AxisPosition == eAxisPosition.Left)
                {
                    foreach (var tb in ret)
                    {
                        tb.Left += (widest - tb.Width);
                    }
                }
                else
                {
                    foreach (var tb in ret)
                    {
                        tb.Left += LeftMargin;
                    }
                }
            }
            else if(LabelOrientation==eTextOrientation.Horizontal) //Only apples when labels are horizontally aligned
            {
                //Align the axis labels according to the label alignment setting. This is only relevant for horizontal axis, vertical axis are always right aligned.
                var lblAlignment = (Axis as ExcelChartAxisStandard)?.LabelAlignment??OfficeOpenXml.eAxisLabelAlignment.Center;
                var majorWidth = Rectangle.Width / AxisValues.Count;
                foreach (var tb in ret)
                {
                    switch (lblAlignment)
                    {
                        case OfficeOpenXml.eAxisLabelAlignment.Left:
                            break;
                        case OfficeOpenXml.eAxisLabelAlignment.Center:
                            tb.Left += majorWidth / 2 - tb.Width / 2;
                            break;
                        case OfficeOpenXml.eAxisLabelAlignment.Right:
                            tb.Left += majorWidth - tb.Width;
                            break;
                    }
                }
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
                if ((Axis.AxisType == eAxisType.Cat || Axis.IsVertical==false) && LabelOrientation==eTextOrientation.Horizontal)
                {
                    //Between tickmarks
                    var majorWidth = Rectangle.Width / AxisValues.Count;
                    var majorTickStartingPosition = Rectangle.Left + majorWidth * i;
                    //var middleOfBounds = majorTickStartingPosition + (majorWidth / 2);
                    return majorTickStartingPosition;
                }
                else
                {
                    if(Axis.AxisType == eAxisType.Cat)
                    {
                        var majorWidth = Rectangle.Width / AxisValues.Count;
                        var majorTickStartingPosition = Rectangle.Left + majorWidth * i;
                        //var middleOfBounds = majorTickStartingPosition + (majorWidth / 2);
                        return majorTickStartingPosition;
                    }
                    else
                    {
                        var min = ConvertUtil.GetValueDouble(Values[0]);
                        var max = ConvertUtil.GetValueDouble(Values.Last());
                        var v = ConvertUtil.GetValueDouble(Values[i]);
                        var majorWidth = Rectangle.Width * (v - Min) / (Max - Min);
                        return Rectangle.Left + majorWidth;
                    }
                }
            }
        }

        private double GetAxisItemTop(int i, OfficeOpenXml.Interfaces.Drawing.Text.TextMeasurement m)
        {
            if (Axis.AxisPosition == eAxisPosition.Top)
            {
                switch (LabelOrientation)
                {
                    case eTextOrientation.Vertical:
                    case eTextOrientation.Diagonal:
                        return Rectangle.Bottom;
                    default:
                        return Rectangle.Bottom - m.Height - TopMargin;
                }
            }
            else if (Axis.AxisPosition == eAxisPosition.Bottom)
            {
                switch(LabelOrientation)
                {
                    case eTextOrientation.Vertical:
                        if (Axis.LabelPosition == eTickLabelPosition.Low)
                        {
                            return Rectangle.Bottom - m.Width;
                        }
                        else
                        {
                            return Rectangle.Top;
                        }
                    case eTextOrientation.Diagonal:
                        if (Axis.LabelPosition == eTickLabelPosition.Low)
                        {
                            return Rectangle.Bottom - (m.Width+m.Height)* COS45;
                        }
                        else
                        {
                            return Rectangle.Top;
                        }
                    default:
                        if (Axis.LabelPosition == eTickLabelPosition.Low)
                        {
                            return Rectangle.Bottom - m.Height;
                        }
                        else if (Axis.LabelPosition == eTickLabelPosition.NextTo)
                        {
                            return Rectangle.Top + BottomMargin;
                        }
                        else //TODO:Add support for hight.
                        {
                            return Rectangle.Top + BottomMargin;
                        }
                }
            }
            else
            {
                var majorHeight = Rectangle.Height / (AxisValues.Count-1);
                if (Axis.AxisType == eAxisType.Cat)
                {
                    return Rectangle.Top + majorHeight * (AxisValues.Count - i - 1) + (majorHeight / 2) - m.Height / 2;
                }
                else
                {
                    //return Rectangle.Top + majorHeight * (AxisValues.Count - i - 1) - m.Height / 2;
                    return Rectangle.Top + majorHeight * (AxisValues.Count - i - 1);
                }
            }

        }

        private List<SvgRenderLineItem> AddTickmarks(double units, eTimeUnit? dateUnit, double parentUnit, double tickMarkWidth, eAxisTickMark type)
        {
            var axisStyle = GetAxisStyleEntry();

            var tms = new List<SvgRenderLineItem>();
            double min, max, addMinor=0D;
            if(double.IsNaN(parentUnit)==false && parentUnit==units)
            {
                addMinor = parentUnit / 2;
            }

            if (Axis.AxisType == eAxisType.Cat)
            {
                min = 1;
                max = AxisValues.Count;
            }
            else
            {
                min = Min;
                max = Max;
            }

            double tickMarkWidthInside=0, tickMarkWidthOutside=0;
            if(type==eAxisTickMark.In || type==eAxisTickMark.Cross)
            {
                tickMarkWidthInside = tickMarkWidth;
            }
            if(type==eAxisTickMark.Out || type == eAxisTickMark.Cross)
            {
                tickMarkWidthOutside = tickMarkWidth;
            }
            var diff = max - min + 1;
            double d = min + addMinor;
            while (d <= max+1)
            {
                if (double.IsNaN(parentUnit) || (d % parentUnit != 0))
                {
                    double x1, y1, x2, y2;
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
                            y1 = (float)Rectangle.Bottom - tickMarkWidthOutside;
                            y2 = (float)Rectangle.Bottom + tickMarkWidthInside;
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
                    if(tm.BorderWidth<1) //Excel seems to have this as minimum width for tick marks, so we enforce it here to make sure they are visible.
                    {
                        tm.BorderWidth = 1;
                    }
                    tms.Add(tm);
                }
                switch (dateUnit)
                {
                    case eTimeUnit.Years:
                        d = DateTime.FromOADate(d).AddYears((int)units).ToOADate();
                        break;
                    case eTimeUnit.Months:
                        if (units>=1D)
                        {
                            d = DateTime.FromOADate(d).AddMonths((int)units).ToOADate();
                        }
                        else
                        {
                            var dt = DateTime.FromOADate(d);
                            var days = DateTime.DaysInMonth(dt.Year, dt.Month) * units;
                            d += days;
                        }
                        break;
                    default:
                        d += units;
                        break;
                }
            }
            return tms;
        }
        private List<RenderItem> AddGridlines(double units, double parentUnit, ExcelDrawingBorder lineItem, ExcelChartStyleEntry styleEntry)
        {
            var axisStyle = GetAxisStyleEntry();

            var tms = new List<RenderItem>();
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

            List<Point> points = new List<Point>();

            for (double d = min; d <= Max; d += units)
            {
                if(d==min && Line!=null && Line.BorderWidth>0) continue;
                if (double.IsNaN(parentUnit) || (d % parentUnit != 0))
                {
                    //float x1, y1, x2, y2;
                    switch (Axis.AxisPosition)
                    {
                        case eAxisPosition.Left:
                        case eAxisPosition.Right:
                            points.Add(new Point(0f, (float)(pa.Rectangle.Top + pa.Rectangle.Height - ((d - min) / diff * pa.Rectangle.Height))));
                            //y1 = (float)(pa.Rectangle.Top + pa.Rectangle.Height - ((d - min) / diff * pa.Rectangle.Height));
                            //y2 = y1;
                            //x1 = (float)pa.Rectangle.Right;
                            //x2 = (float)pa.Rectangle.Left;
                            break;
                        case eAxisPosition.Top:
                        case eAxisPosition.Bottom:
                            points.Add(new Point((float)(pa.Rectangle.Left + ((d - min) / diff * pa.Rectangle.Width)), 0f));
                            //x1 = (float)(pa.Rectangle.Left + ((d - min) / diff * pa.Rectangle.Width));
                            //x2 = x1;
                            //y1 = (float)pa.Rectangle.Top;
                            //y2 = (float)pa.Rectangle.Bottom;
                            break;
                        default:
                            throw new InvalidOperationException("Invalid axis position");
                    }

                    //var tm = new SvgRenderLineItem(SvgChart, SvgChart.Bounds);
                    //tm.X1 = x1;
                    //tm.Y1 = y1;
                    //tm.X2 = x2;
                    //tm.Y2 = y2;
                    //tm.SetDrawingPropertiesBorder(lineItem, styleEntry.BorderReference.Color, true, lineItem.Width);
                    //tms.Add(tm);
                }
            }

            float x1, y1, x2, y2;

            string id = "";

            switch (Axis.AxisPosition)
            {
                case eAxisPosition.Left:
                case eAxisPosition.Right:
                    id = "xGridLine";
                    y1 = (float)points.Last().Top;
                    y2 = y1;
                    x1 = (float)pa.Rectangle.Right;
                    x2 = (float)pa.Rectangle.Left;
                    break;
                case eAxisPosition.Top:
                case eAxisPosition.Bottom:
                    id = "yGridLine";
                    x1 = (float)points.Last().Left;
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

            tm.DefId = id;

            tms.Add(tm);

            var distY = points[0].Top - (float)points.Last().Top;
            var distX = points[0].Left - (float)points.Last().Left;

            var offsetX = distX / (points.Count-1);
            var offsetY = distY / (points.Count-1);

            for(int i = 0; i < points.Count; i++)
            {
                var refItem = new SvgUseRefItem(SvgChart, SvgChart.Bounds, id);
                if(id == "xGridLine")
                {
                    refItem.Y = offsetY*i;
                    refItem.X = 0f;
                }
                else if(id == "yGridLine")
                {
                    refItem.X = offsetX*i;
                    refItem.Y = 0f;
                }
                tms.Add(refItem);
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
                    var diff = Max - Min+1;
                    return (((val-Min) / diff * SvgChart.Plotarea.Rectangle.Width));
                }
            }
        }
        protected List<object> GetAxisValue(ExcelChartAxisStandard ax, RenderItem rect, out double? min, out double? max, out double? majorUnit, out eTimeUnit? dateUnit, out eTextOrientation orientation)
        {
            var values = ax.GetAxisValues(out bool isCount);

            var options = new AxisOptions
            {
                LockedMin = ax.MinValue,
                LockedMax = ax.MaxValue,
                LockedInterval = ax.MajorUnit,
                LockedIntervalUnit = ax.MajorTimeUnit,
                AddPadding = ax.AxisPosition == eAxisPosition.Left || ax.AxisPosition == eAxisPosition.Right,
                Axis = ax,
                IsStacked100 = Chart.IsTypePercentStacked(),
                ChartSize = rect
            };

            if (ax.AxisType == eAxisType.Cat &&
                isCount == false)
            {
                min = 0;
                max = values.Count;
                majorUnit = 1;
                dateUnit = null;
                orientation = eTextOrientation.Horizontal;
                var res = CategoryAxisScaleCalculator.CalculateByWidth(ref values, SvgChart.TextMeasurer, options);
                
                min = res.Min;
                max = res.Max;
                majorUnit = res.MajorInterval;
                dateUnit = null;
                orientation = res.TextOrientation;

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

            var length = ax.AxisPosition == eAxisPosition.Left || ax.AxisPosition == eAxisPosition.Right ? SvgChart.Bounds.Height : SvgChart.Bounds.Width; //Fix and use plotarea width/height.
            if(isCount)
            {
                majorUnit = 1;
                dateUnit = null;
                for (int i=1;i<=max;i++)
                {
                    l.Add(i);
                }
                var res = CategoryAxisScaleCalculator.CalculateByWidth(ref l, SvgChart.TextMeasurer, options);

                min = res.Min;
                max = res.Max;
                majorUnit = res.MajorInterval;
                dateUnit = null;
                orientation = res.TextOrientation;

                return l.ToList();
            }
            if (ax.IsDate)
            {
                AxisScale res;
                if (ax.IsVertical)
                {
                    res = DateAxisScaleCalculator.Calculate(min ?? 0, max ?? 0, length, options);
                }
                else
                {
                    res = DateAxisScaleCalculator.CalculateByWidth(min ?? 0D, max ?? 0D, SvgChart.TextMeasurer, options);
                }
                orientation = res.TextOrientation;
                dateUnit = res.MajorDateUnit;
                var dt = DateTime.FromOADate(res.Min);
                var maxDt = DateTime.FromOADate(res.Max);
                while (dt <= maxDt)
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
                dateUnit= null;
                orientation = eTextOrientation.Horizontal;
            }

            return l;
        }
    }
}