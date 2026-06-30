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
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer;
using EPPlus.Export.ImageRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Svg.Chart.Util;
using EPPlus.Export.Renderer;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;

namespace EPPlusImageRenderer.Svg
{
    internal class ChartAxisRenderer : ChartDrawingObject, IDrawingChartAxis
    {
        private const double COS45 = 0.70710678118654757; //Constant for Math.Sin(Math.PI / 4) --45 degrees
        internal ChartAxisRenderer(ChartRenderer sc, ExcelChartAxisStandard ax) : base(sc)
        {
            Axis = ax;
            SetMargins(ax.TextBody);

            if (sc.Chart.Series.Count == 0)
            {
                return;
            }

            if(ax.HasTitle)
            {
                Title = new ChartTitleRenderer(sc, ax.Title, "Axis Title", this);
            }
            else
            {
                Title = null;
            }

            Values = GetAxisValue(ax, sc.ChartArea.Rectangle, out double? min, out double? max, out double? majorUnit, out eTimeUnit? dateUnit, out eTextOrientation orientation);
            AxisValues = GetAxisDisplayValues(ax, Values, min, max, majorUnit);

            Min = min ?? 0D;
            Max = max ?? (Values.Count > 0 ? ConvertUtil.GetValueDouble(Values[Values.Count - 1], false, true) : 0D);
            MajorUnit = majorUnit ?? 1;
            MinorUnit = ax.MinorUnit ?? GetAutoMinUnit(MajorUnit);
            MajorDateUnit = dateUnit;
            LabelOrientation = orientation;

            if(Rectangle == null)
            {
                Rectangle = new RectRenderItem(sc.Bounds);
            }

            if (ax.Deleted == false)
            {
                if (ax.Layout.HasLayout)
                {
                    Rectangle = GetRectFromManualLayout(sc, ax.Layout);
                }
                else
                {
                    var aav = ax.ActualAxisPosition;
                    if (ax.IsVertical)
                    {
                        if (aav == eActualAxisPosition.Left || 
                            aav == eActualAxisPosition.LeftSecond)
                        {
                            Rectangle.Width = GetTextWidest(ax) + LeftMargin;
                        }
                        else
                        {
                            Rectangle.Width = GetTextWidest(ax) + RightMargin;
                        }
                    }
                    else
                    {
                        Rectangle.Height = GetTextHeight(ax);
                    }
                }

                Rectangle.FillColor = "none";

                Line = new LineRenderItem(Rectangle.Bounds);
                Line.SetDrawingPropertiesBorder(ChartRenderer.Theme, ax.Border, sc.Chart.StyleManager.Style?.Title.BorderReference.Color, ax.Border.Fill.Style != eFillStyle.NoFill, 1);
                if(Line.BorderWidth < 1)
                {
                    Line.BorderWidth = 1;
                }
            }
        }

        private double GetAutoMinUnit(double majorUnit)
        {
            return majorUnit / 5;
        }

        internal ExcelChartAxisStandard Axis { get; }
        internal LineRenderItem Line { get; set; }
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
        private double GetTextHeight(ExcelChartAxisStandard ax)
        {
            var tm = ChartRenderer.TextMeasurer;
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
                        var width = (m.Width) * COS45;
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

        private double GetTextWidest(ExcelChartAxisStandard ax)
        {
            var mf = ax.Font.GetMeasureFont();
            var shaper = RenderContext.FontEngine.GetShaperForFont(mf);
            var tm = new OpenTypeFontTextMeasurer(shaper);
            
            var widest = 0f;

            foreach(var s in AxisValues)
            {
                var m= tm.MeasureText(s, mf);
                if (m.Width > widest)
                {
                    widest = m.Width;
                }
            }
            return widest;
        }

        public List<object> Values
        {
            get;
            private set;
        }
        public List<string> AxisValues { get; private set; }

        public List<LineRenderItem> MajorTickMarkPositions { get; private set; }
        public List<LineRenderItem> MinorTickMarkPositions { get; private set; }
        public List<RenderItem> MajorGridlinePositions { get; private set; }
        public List<RenderItem> MinorGridlinePositions { get; private set; }
        public ChartAxisTextBoxes  Textboxes{get; private set;}
        public ChartTitleRenderer Title { get; set; }
        public double Min { get; set; }
        public double Max { get; set; }
        public double MajorUnit { get; set; }
        public double MinorUnit { get; set; }
        public eTimeUnit? MajorDateUnit { get; set; }
        public eTextOrientation LabelOrientation { get; set; }
        public bool IsDateScale
        {
            get;
            private set;
        } = false;

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            Title?.AppendRenderItems(renderItems);
            //Title?.Render(sb);
            if(Rectangle!=null || Rectangle.Width==0 || Rectangle.Height==0) renderItems.Add(Rectangle);

            var plotareaGroup = ChartRenderer.Plotarea.Group;
            if (MajorGridlinePositions != null)
            {
                foreach (var tm in MajorGridlinePositions)
                {
                    plotareaGroup.RenderItems.Add(tm);
                }
            }
            if (MinorGridlinePositions != null)
            {
                foreach (var tm in MinorGridlinePositions)
                {
                    plotareaGroup.RenderItems.Add(tm);
                }
            }

            if (Line != null) renderItems.Add(Line);

            if (MajorTickMarkPositions != null)
            {
                foreach (var tm in MajorTickMarkPositions)
                {
                    renderItems.Add(tm);
                }
            }
            if (MinorTickMarkPositions != null)
            {
                foreach (var tm in MinorTickMarkPositions)
                {
                    renderItems.Add(tm);
                }
            }
            
            //The axis text boxes is rendered later as they have a higher Z-order.
        }


        internal void AddTickmarksAndValues(List<RenderItem> DefItems)
        {
            if (Axis.Deleted == true) return;
            if (Axis.MajorTickMark != eAxisTickMark.None)
            {
                MajorTickMarkPositions = AddTickmarks(MajorUnit, MajorDateUnit, double.NaN, 4D.PixelToPoint(), Axis.MajorTickMark);
            }

            if (Axis.MinorTickMark != eAxisTickMark.None)
            {
                MinorTickMarkPositions = AddTickmarks(MinorUnit, MajorDateUnit, MajorUnit, 2D.PixelToPoint(), Axis.MinorTickMark);
            }

            if(Axis.HasMajorGridlines)
            {
                MajorGridlinePositions = AddGridlines(MajorUnit, double.NaN, Axis.MajorGridlines, Chart.StyleManager.Style?.GridlineMajor);
            }

            if ((Axis.HasMinorGridlines))
            {
                MinorGridlinePositions = AddGridlines(MinorUnit, MajorUnit, Axis.MinorGridlines, Chart.StyleManager.Style.GridlineMinor);
            }

            if (AxisValues != null && AxisValues.Count > 0 && Axis.Deleted==false && Axis.LabelPosition != eTickLabelPosition.None)
            {
                Textboxes = new ChartAxisTextBoxes(ChartRenderer);
                Textboxes.TextBoxes = GetAxisValueTextBoxes();  
            }
        }

        private List<DrawingTextBox> GetAxisValueTextBoxes()
        {
            var ret = new List<DrawingTextBox>();
            if (Axis.LabelPosition == eTickLabelPosition.None) return ret;

            var mf = Axis.Font.GetMeasureFont();

            var shaper = RenderContext.FontEngine.GetShaperForFont(mf);
            var tm = new OpenTypeFontTextMeasurer(shaper);

            var axisStyle = GetAxisStyleEntry();
            double maxWidth, maxHeight;
            if(Axis.AxisPosition==eAxisPosition.Left || Axis.AxisPosition == eAxisPosition.Right)
            {
                maxWidth = ChartRenderer.ChartArea.Rectangle.Width / 3; //TODO: Check this value.
                maxHeight = Rectangle.Height / AxisValues.Count;
            }
            else
            {
                switch (LabelOrientation)
                {
                    case eTextOrientation.Vertical:
                        maxWidth = ChartRenderer.ChartArea.Rectangle.Height / 3;
                        maxHeight = Rectangle.Width / AxisValues.Count; //TODO: Check this value.
                        break;                    
                    case eTextOrientation.Diagonal:
                        maxWidth = (Rectangle.Width + Rectangle.Height) / COS45;
                        maxHeight = ChartRenderer.ChartArea.Rectangle.Height / 3; //TODO: Check this value.
                        break;
                    default:
                        maxWidth = Rectangle.Width / AxisValues.Count;
                        maxHeight = ChartRenderer.ChartArea.Rectangle.Height / 3; //TODO: Check this value.
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
                    if (Axis.AxisType == eAxisType.Cat || Axis.AxisType==eAxisType.Date)
                    {
                        x = ticMarkX;
                        y = ticMarkY;
                    }
                    else
                    {
                        if(Axis.IsVertical)
                        {
                            x = ticMarkX;
                            if (ChartRenderer.Chart.IsTypeBar())
                            {
                                y = ticMarkY;
                            }
                            else
                            {
                                y = ticMarkY - height / 2;
                            }
                        }
                        else
                        {
                            x = ticMarkX  - width / 2;
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
                        if (Axis.ActualAxisPosition == eActualAxisPosition.Bottom || Axis.ActualAxisPosition == eActualAxisPosition.BottomSecond)
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
                        if (Axis.ActualAxisPosition == eActualAxisPosition.Bottom || Axis.ActualAxisPosition == eActualAxisPosition.BottomSecond)
                        {
                            y = ticMarkY + BottomMargin + 4;
                        }
                        else //Top
                        {
                            y = ticMarkY - TopMargin - 4;
                        }
                    }
                }
                
                var tb = new DrawingTextBox(Chart, Rectangle.Bounds, x, y, width, height, maxWidth, maxHeight);
                if (LabelOrientation == eTextOrientation.Diagonal)
                {
                    tb.Rotation = -45;
                    if(Axis.ActualAxisPosition==eActualAxisPosition.Bottom || Axis.ActualAxisPosition == eActualAxisPosition.BottomSecond)
                    {
                        tb.TextAnchor = eTextAnchor.End;
                    }
                }

                else if (LabelOrientation == eTextOrientation.Vertical)
                {
                    tb.Rotation = -90;
                    if (Axis.ActualAxisPosition == eActualAxisPosition.Bottom || Axis.ActualAxisPosition == eActualAxisPosition.BottomSecond)
                    {
                        tb.TextAnchor = eTextAnchor.End;
                    }
                }

                var p = Axis.TextBody.Paragraphs.FirstOrDefault();

                if (p.HorizontalAlignment != eTextAlignment.Center && Axis.AxisType!=eAxisType.Val && (Axis.AxisPosition == eAxisPosition.Bottom || Axis.AxisPosition == eAxisPosition.Top))
                {
                    //Horizontal axises are always center aligned visually
                    //Should be broken out as input to ImportParagraph instead of changing the base item
                    p.HorizontalAlignment = eTextAlignment.Center;
                }

                tb.ImportParagraph(p, 0, v);

                //tb.TextBody.Paragraphs[0].AddText(v, Axis.Font);
                tb.Rectangle.SetDrawingPropertiesFill(ChartRenderer.Theme, Axis.Fill, axisStyle?.FillReference.Color, true);

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
            else if(LabelOrientation==eTextOrientation.Horizontal && IsCatAx()) //Only apples when labels are horizontally aligned
            {
                //Align the axis labels according to the label alignment setting. This is only relevant for horizontal axis, vertical axis are always right aligned.
                var lblAlignment = (Axis as ExcelChartAxisStandard)?.LabelAlignment ?? OfficeOpenXml.eAxisLabelAlignment.Center;
                var majorWidth = Rectangle.Width / AxisValues.Count;
                if (Axis.CrossingAxis == null || Axis.CrossingAxis.CrossBetween == eCrossBetween.MidCat)
                {
                    foreach (var tb in ret)
                    {
                        switch (lblAlignment)
                        {
                            case OfficeOpenXml.eAxisLabelAlignment.Left:
                                tb.Left -= (tb.Width + majorWidth) / 2;
                                break;
                            case OfficeOpenXml.eAxisLabelAlignment.Center:
                                tb.Left -= tb.Width / 2;
                                break;
                            case OfficeOpenXml.eAxisLabelAlignment.Right:
                                tb.Left += (tb.Width + majorWidth) / 2;
                                break;
                        }
                    }
                }
                else
                {
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
            }
            else if(LabelOrientation == eTextOrientation.Diagonal)
            {
                if (!(Axis.CrossingAxis == null || Axis.CrossingAxis.CrossBetween == eCrossBetween.MidCat))
                {
                    var majorWidth = Rectangle.Width / AxisValues.Count;
                    foreach (var tb in ret)
                    {
                        tb.Left += majorWidth / 2;
                    }
                }
            }

            return ret;
        }

        private bool IsCatAx()
        {
            return Axis.AxisType == eAxisType.Cat || (Axis.AxisType == eAxisType.Date && IsDateScale==false);
        }

        private double GetAxisItemLeft(int i, OfficeOpenXml.Interfaces.Drawing.Text.TextMeasurement m)
        {
            if (Axis.IsVertical)
            {
                return Rectangle.Left;
            }
            else
            {
                if (IsCatAx())
                {
                    double majorWidth;
                    if (Axis.CrossingAxis == null || Axis.CrossingAxis.CrossBetween == eCrossBetween.Between)
                    {
                        majorWidth = Rectangle.Width / AxisValues.Count;
                    }
                    else
                    {
                        majorWidth = Rectangle.Width / (AxisValues.Count - 1);
                    }
                    var majorTickStartingPosition = Rectangle.Left + majorWidth * i;
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
                //}
            }
        }

        private double GetAxisItemTop(int i, OfficeOpenXml.Interfaces.Drawing.Text.TextMeasurement m)
        {
            if (Axis.ActualAxisPosition == eActualAxisPosition.Top || Axis.ActualAxisPosition == eActualAxisPosition.TopSecond)
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
            else if (Axis.ActualAxisPosition == eActualAxisPosition.Bottom || Axis.ActualAxisPosition == eActualAxisPosition.BottomSecond)
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
                if (Axis.AxisType == eAxisType.Cat || Axis.AxisType == eAxisType.Date)
                {
                    var majorHeight = Rectangle.Height / (AxisValues.Count);
                    return Rectangle.Top + majorHeight * (AxisValues.Count - i) - ((majorHeight / 2) + m.Height / 2);
                }
                else
                {
                    var majorHeight = Rectangle.Height / (AxisValues.Count-1);
                    return Rectangle.Top + majorHeight * (AxisValues.Count - i - 1);
                    //return Rectangle.Top + majorHeight * (AxisValues.Count - i - 1);
                }
            }

        }

        private List<LineRenderItem> AddTickmarks(double units, eTimeUnit? dateUnit, double parentUnit, double tickMarkWidth, eAxisTickMark type)
        {
            var axisStyle = GetAxisStyleEntry();

            var tms = new List<LineRenderItem>();
            double min, max, addMinor=0D;
            if(double.IsNaN(parentUnit)==false && parentUnit==units)
            {
                addMinor = parentUnit / 2;
            }

            if (Axis.AxisType == eAxisType.Cat)
            {
                min = 1;
                if (Axis.CrossingAxis==null || Axis.CrossingAxis.CrossBetween == eCrossBetween.Between)
                {
                    max = AxisValues.Count;
                }
                else
                {
                    max = AxisValues.Count - 1;
                }
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
                    switch (Axis.ActualAxisPosition)
                    {
                        case eActualAxisPosition.Left:
                        case eActualAxisPosition.LeftSecond:
                            y1 = (float)(Rectangle.Top + Rectangle.Height - ((d - min) / diff * Rectangle.Height));
                            y2 = y1;                            
                            x1 = (float)Rectangle.Right - tickMarkWidthOutside;
                            x2 = (float)Rectangle.Right + tickMarkWidthInside;
                            break;
                        case eActualAxisPosition.Right:
                        case eActualAxisPosition.RightSecond:
                            y1 = (float)(Rectangle.Top + Rectangle.Height - ((d - min) / diff * Rectangle.Height));
                            y2 = y1;
                            x1 = (float)Rectangle.Left - tickMarkWidthInside;
                            x2 = (float)Rectangle.Left + tickMarkWidthOutside;
                            break;
                        case eActualAxisPosition.Top:
                        case eActualAxisPosition.TopSecond:
                            x1 = (float)(Rectangle.Left + ((d - min) / diff * Rectangle.Width));
                            x2 = x1;
                            y1 = (float)Rectangle.Bottom - tickMarkWidthOutside;
                            y2 = (float)Rectangle.Bottom + tickMarkWidthInside;
                            break;
                        case eActualAxisPosition.Bottom:
                        case eActualAxisPosition.BottomSecond:
                            x1 = (float)(Rectangle.Left + ((d - min) / diff * Rectangle.Width));
                            x2 = x1;
                            y1 = (float)Rectangle.Top - tickMarkWidthInside;
                            y2 = (float)Rectangle.Top + tickMarkWidthOutside;
                            break;
                        default:
                            throw new InvalidOperationException("Invalid axis position");
                    }
                    var tm = new LineRenderItem(ChartRenderer.Bounds);
                    tm.X1 = x1;
                    tm.Y1 = y1;
                    tm.X2 = x2;
                    tm.Y2 = y2;
                    tm.SetDrawingPropertiesBorder(ChartRenderer.Theme, Axis.Border, axisStyle?.BorderReference.Color, true);
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
            var pa = ChartRenderer.Plotarea;
            var diff = Max - min;

            List<Point> points = new List<Point>();
            var group = ChartRenderer.Plotarea.Group;
            for (double d = min; d <= Max; d += units)
            {
                if(d==min && Line!=null && Line.BorderWidth>0) continue;
                if (double.IsNaN(parentUnit) || (d % parentUnit != 0))
                {
                    switch (Axis.AxisPosition)
                    {
                        case eAxisPosition.Left:
                        case eAxisPosition.Right:
                            points.Add(new Point(0f, (float)(pa.Rectangle.Height - ((d - min) / diff * pa.Rectangle.Height))));
                            break;
                        case eAxisPosition.Top:
                        case eAxisPosition.Bottom:
                            var xValue = (float)(((d - min) / diff * pa.Rectangle.Width));
                            points.Add(new Point(xValue, 0f));
                            break;
                        default:
                            throw new InvalidOperationException("Invalid axis position");
                    }
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
                    x1 = (float)0;
                    x2 = (float)pa.Rectangle.Width;
                    break;
                case eAxisPosition.Top:
                case eAxisPosition.Bottom:
                    id = "yGridLine";
                    x1 = (float)points[0].Left;
                    x2 = x1;
                    y1 = 0;
                    y2 = (float)pa.Rectangle.Height;
                    break;
                default:
                    throw new InvalidOperationException("Invalid axis position");
            }

            var tm = new LineRenderItem(ChartRenderer.Bounds);
            tm.X1 = x1;
            tm.Y1 = y1;
            tm.X2 = x2;
            tm.Y2 = y2;
            tm.SetDrawingPropertiesBorder(ChartRenderer.Theme, lineItem, styleEntry?.BorderReference.Color, true, lineItem.Width);

            tm.DefId = id;

            tms.Add(tm);

            var distX = (float)points.Last().Left - points[0].Left;
            var distY = points[0].Top - (float)points.Last().Top;

            var offsetX = distX / (points.Count-1);
            var offsetY = distY / (points.Count-1);

            for(int i = 0; i < points.Count; i++)
            {
                var refItem = new UseReferenceRenderItem(ChartRenderer.Bounds, "#" + id);
                if(id == "xGridLine")
                {
                    refItem.X = 0f;
                    refItem.Y = offsetY*i;
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
                    axisStyle = Chart.StyleManager.Style?.CategoryAxis;
                    break;
                case eAxisType.Serie:
                    axisStyle = Chart.StyleManager.Style?.SeriesAxis;
                    break;
                default:
                    axisStyle = Chart.StyleManager.Style?.ValueAxis;
                    break;
            }

            return axisStyle;
        }

        internal double GetPositionInPlotarea(double val, bool startValue=false)
        {
            if (Axis.AxisPosition == eAxisPosition.Left || Axis.AxisPosition == eAxisPosition.Right)
            {
                if (Axis.AxisType == eAxisType.Cat)
                {
                    var majorHeight = ChartRenderer.Plotarea.Rectangle.Height / Max;
                    if(startValue)
                    {
                        return majorHeight * val;
                    }
                    else
                    {
                        //if (Axis.CrossingAxis == null || Axis.CrossingAxis.CrossBetween == eCrossBetween.Between)
                        //{
                        //    majorHeight = DrawingChart.Plotarea.Rectangle.Height / (Max - 1);
                        //    return majorHeight * val;
                        //}
                        //else
                        //{
                            return majorHeight * val + (majorHeight / 2);
                        //}
                    }
                }
                else if (Axis.AxisType == eAxisType.Date && IsDateScale == false)
                {
                    //if (val < Min || val > Max) return double.NaN;
                    var diff = Max - Min + 1;
                    return (((val - Min) / diff * ChartRenderer.Plotarea.Rectangle.Height));
                }
                else
                {
                    //if (val < Min || val > Max) return double.NaN;
                    var diff = Max - Min;
                    return (Max - val) / diff * ChartRenderer.Plotarea.Rectangle.Height;
                }
            }
            else
            {
                if (Axis.AxisType == eAxisType.Cat)
                {
                    var majorWidth = ChartRenderer.Plotarea.Rectangle.Width / Max;
                    if (startValue)
                    {
                        return majorWidth * val;
                    }
                    else
                    {
                        if(Axis.CrossingAxis == null || Axis.CrossingAxis.CrossBetween==eCrossBetween.Between)
                        {
                            return majorWidth * val + (majorWidth / 2);
                        }
                        else
                        {
                            majorWidth = ChartRenderer.Plotarea.Rectangle.Width / (Max - 1);
                            return majorWidth * val;
                        }
                    }
                }
                else if(Axis.AxisType == eAxisType.Date && IsDateScale==false)
                {
                    if (val < Min || val > Max) return double.NaN;
                    var diff = Max - Min + 1;
                    return (((val - Min) / diff * ChartRenderer.Plotarea.Rectangle.Width));
                }
                else
                {
                    if (val < Min || val > Max) return double.NaN;
                    var diff = Max - Min;
                    return (((val - Min) / diff * ChartRenderer.Plotarea.Rectangle.Width));
                }
            }
        }
        protected List<object> GetAxisValue(ExcelChartAxisStandard ax, RenderItem rect, out double? min, out double? max, out double? majorUnit, out eTimeUnit? dateUnit, out eTextOrientation orientation)
        {
            var values = ax.GetAxisValues(out bool isCount);
            var isNumeric = values.Any(x => x == null || x.IsNumeric());
            var options = new AxisOptions
            {
                LockedMin = ax.MinValue,
                LockedMax = ax.MaxValue,
                LockedInterval = ax.MajorUnit,
                LockedIntervalUnit = ax.MajorTimeUnit,
                AddPadding = ShouldHavePadding(),
                Axis = ax,
                IsStacked100 = Chart.IsTypePercentStacked(),
                ChartSize = rect
            };

            if ((ax.AxisType == eAxisType.Cat || (ax.IsDate && isNumeric==false)) &&
                isCount == false)
            {
                AxisScale res;
                if (ax.IsVertical)
                {
                    res = CategoryAxisScaleCalculator.CalculateVerticalAxisByHeight(ref values, ChartRenderer.TextMeasurer, options);
                }
                else
                {
                    res = CategoryAxisScaleCalculator.CalculateHorizontalAxisByWidth(ref values, ChartRenderer.TextMeasurer, options);
                }

                min = res.Min;
                max = res.Max;
                majorUnit = res.MajorInterval;
                dateUnit = null;
                orientation = res.TextOrientation;                

                return res.DisplayValues;
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

            var length = ax.AxisPosition == eAxisPosition.Left || ax.AxisPosition == eAxisPosition.Right ? ChartRenderer.Bounds.Height : ChartRenderer.Bounds.Width; //Fix and use plotarea width/height.
            if(isCount)
            {
                majorUnit = 1;
                dateUnit = null;
                for (int i=1;i<=max;i++)
                {
                    l.Add(i);
                }
                var res = CategoryAxisScaleCalculator.CalculateHorizontalAxisByWidth(ref l, ChartRenderer.TextMeasurer, options);

                min = res.Min;
                max = res.Max;
                majorUnit = res.MajorInterval;
                dateUnit = null;
                orientation = res.TextOrientation;

                return l.ToList();
            }
            if(ax.AxisType==eAxisType.Val)
            {
                AdjustminMaxFromChartObjects(ax, ref min, ref max);
            }
            if (ax.IsDate)
            {
                AxisScale res;
                if (ax.IsVertical)
                {
                    res = DateAxisScaleCalculator.CalculateByWidthHeight(options.ChartSize.Bounds.Height, min ?? 0D, max ?? 0D, ChartRenderer.TextMeasurer, options);
                }
                else
                {
                    if (ax.AxisType==eAxisType.Val)
                    {
                        res = DateAxisScaleCalculator.Calculate(min ?? 0D, max ?? 0D, options);
                    }
                    else
                    {
                        res = DateAxisScaleCalculator.CalculateByWidthAllowDiagonal(min ?? 0D, max ?? 0D, ChartRenderer.TextMeasurer, options);
                    }
                }

                orientation = res.TextOrientation;
                dateUnit = res.MajorDateUnit;
                majorUnit = res.MajorInterval;
                var dt = DateTime.FromOADate(res.Min);
                var maxDt = DateTime.FromOADate(res.Max);
                IsDateScale = (dateUnit != eTimeUnit.Days || majorUnit > 1) && values.Count > 31;

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

        /// <summary>
        /// Adjust the min and max values based on the values in the chart objects, such as trendlines. This is needed to make sure that the trendlines are visible in the chart and not cut off because the axis scale is based on the data series only.
        /// </summary>
        /// <param name="min">The min value to adjust.</param>
        /// <param name="max">The max value to adjust.</param>
        private void AdjustminMaxFromChartObjects(ExcelChartAxisStandard ax, ref double? min, ref double? max)
        {
            foreach (var drawer in ChartRenderer.Plotarea.ChartTypeDrawers)
            {
                if (drawer.IsOnAxis(ax))
                {
                    if (drawer.SupportsTrendlines)
                    {
                        foreach (var tl in drawer.Trendlines)
                        {
                            foreach (var c in tl.Coordinates)
                            {
                                if (min > c.Y)
                                {
                                    min = c.Y;
                                }
                                if (max < c.Y)
                                {
                                    max = c.Y;
                                }
                            }
                        }
                    }
                    if(drawer.SupportsErrorBars && drawer.ErrorBars!=null)
                    {
                        foreach(var v in drawer.ErrorBars.Values)
                        {
                            if (v[0] < min)
                            {
                                min = v[0];
                            }
                            if (v[2] > max)
                            {
                                max = v[2];
                            }
                        }

                    }
                }
            }
        }

        private bool ShouldHavePadding()
        {
            return Axis.AxisType == eAxisType.Val || (Chart.IsTypeLine() && Axis.AxisType == eAxisType.Date);
        }
    }
}