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
using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Svg.Chart.Util;
using EPPlus.Export.Renderer;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Renderer.Chart;
using OfficeOpenXml.Drawing.Renderer.Chart.Defaults;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace EPPlusImageRenderer.Svg
{
    internal class ChartAxisRenderer : ChartDrawingObjectWithBackground, IDrawingChartAxis
    {
        private const double COS45 = 0.70710678118654757; //Constant for Math.Sin(Math.PI / 4) --45 degrees

        internal override System.Drawing.Color? DefaultFillColor { get; }
        int _axisMaxItems=-1;
        internal int AxisMaxItems 
        {
            get
            {
                if (_axisMaxItems<0)
                {
                    var mf = Axis.Font.GetMeasureFont(DefaultFonts2007Style.Axis);
                    var shaper = ChartRenderer.RenderContext.GetMeasurementShaper(new FontFormatBase(mf));
                    if (Axis.IsVertical)
                    {
                        var lineHeight = shaper.GetLineHeightInPoints(mf.Size);
                        _axisMaxItems = (int)(ChartRenderer.Plotarea.Rectangle.Height / lineHeight * 0.8);
                    }
                    else
                    {
                        var lineWidth = mf.Size*5; //Estimate the width of a character in points. This is a rough estimate and may not be accurate for all fonts.
                        _axisMaxItems = (int)(Rectangle.Width / GetTextWidest(Axis) * 1.3);
                    }
                    //_axisMaxItems -= _axisMaxItems % 2; //Removed for now
                }
                return _axisMaxItems;
            }
        }
        internal ChartAxisRenderer(ChartRenderer sc, ExcelChartAxisStandard ax) : base(sc)
        {
            Axis = ax;
            SetMargins(ax.TextBody);

            DefaultFillColor = System.Drawing.Color.Transparent;

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
            if(AutoAxisType == eAxisType.Cat && !(IsDateAutoAxis || IsDateScale))
            {
                MinorUnit = (ax.MinorUnit ?? MajorUnit) / 2D;
            }
            else if(IsDateAutoAxis || IsDateScale)
            {
                MinorUnit = ax.MinorUnit ?? 1;
            }
            else
            {
                MinorUnit = ax.MinorUnit ?? GetAutoMinUnit(MajorUnit);
            }
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
                        if (aav == eAxisPosition.Left)
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

                Rectangle.Style.FillColor = "none";

                Line = new LineRenderItem(Rectangle);
                Line.Style.SetDrawingPropertiesBorder(ChartRenderer.Theme, ax.Border, sc.Chart.StyleManager.Style?.Title.BorderReference.Color, ax.Border.IsEmpty==true || ax.Border.Fill.Style != eFillStyle.NoFill, GetDefaultBorderColor, 1);
                if(Line.Style.BorderWidth < 1)
                {
                    Line.Style.BorderWidth = 1;
                }
            }
        }
        public eAxisType AutoAxisType
        {
            get
            {
                if(Axis.AxisType == eAxisType.Cat)
                {
                    if(IsDateAutoAxis)
                    {
                        return eAxisType.Date;
                    }
                }
                return Axis.AxisType;
            }
        }
        private double GetAutoMinUnit(double majorUnit)
        {
            if (Axis.LogBase.HasValue)
            {
                return 1;
            }
            else
            {
                return majorUnit / 5;
            }
        }

        internal ExcelChartAxisStandard Axis { get; }
        internal LineRenderItem Line { get; set; }
        private List<string> GetAxisDisplayValues(ExcelChartAxisStandard ax, List<object> values, double? min, double? max, double? majorUnit)
        {
            var displayValues = new List<string>();
            var format = ax.FormatOrFirstValueFormat;
            var nf = new ExcelFormatTranslator(format, 0);
            //Excel replaces the format with a default date format if the axis is date based.
            if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
            {
                if(format == "m/d/yyyy")
                {
                    var sdFormat = ExcelNumberFormat.GetFromBuildInFromID(14); //14 is standard regional short date.
                    nf = new ExcelFormatTranslator(sdFormat, 14);
                }
            }            
            foreach (var v in CategoryAxisScaleCalculator.GetUniqueValues(values))
            {
                string s;
                if (ax.DisplayUnit > 1)
                {
                    var d = ConvertUtil.GetValueDouble(v, false, true);
                    d /= ax.DisplayUnit;
                    s = ValueToTextHandler.FormatValue(d, false, nf, null, out bool isValidFormat);
                }
                else
                {
                    s = ValueToTextHandler.FormatValue(v, false, nf, null, out bool isValidFormat);
                }
                displayValues.Add(s);
            }
            return displayValues;
        }
        private double GetTextHeight(ExcelChartAxisStandard ax)
        {
            //The shaper is obtained through the render context so the family is substituted
            //for the render target, matching the font used in the output.
            var mf = ax.Font.GetMeasureFont(DefaultFonts2007Style.Axis);
            var shaper = ChartRenderer.RenderContext.GetMeasurementShaper(new FontFormatBase(mf));
            var fontSize = mf.Size;
            var lineHeight = shaper.GetLineHeightInPoints(fontSize);

            var highest = 0D;
            foreach (var s in AxisValues)
            {
                //An empty label has no size, as with ITextMeasurer.
                if (string.IsNullOrEmpty(s))
                    continue;

                switch (LabelOrientation)
                {
                    case eTextOrientation.Horizontal:
                        if (lineHeight > highest)
                        {
                            highest = lineHeight;
                        }
                        break;
                    case eTextOrientation.Custom:
                    case eTextOrientation.Diagonal:
                        var diagonalHeight = shaper.Shape(s).GetWidthInPoints(fontSize) * COS45;
                        if (diagonalHeight > highest)
                        {
                            highest = diagonalHeight;
                        }
                        break;
                    case eTextOrientation.Vertical:
                        var verticalHeight = shaper.Shape(s).GetWidthInPoints(fontSize);
                        if (verticalHeight > highest)
                        {
                            highest = verticalHeight;
                        }
                        break;
                }
            }
            return highest.PointToPixel();
        }

        private double GetTextWidest(ExcelChartAxisStandard ax)
        {
            var mf = ax.Font.GetMeasureFont(DefaultFonts2007Style.Axis);
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
        //internal GetAxisWidthForPlotarea()
        //{

        //}
        //internal double GetAxisHeightForPlotarea()
        //{
        //    var pa = ChartRenderer.Plotarea;

        //    var 
        //}
        public List<object> Values
        {
            get;
            private set;
        }
        public List<string> AxisValues { get; private set; }

        public List<LineRenderItem> MajorTickMarkPositions { get; private set; }
        public List<LineRenderItem> MinorTickMarkPositions { get; private set; }
        public ChartAxisTextBoxes  Textboxes{get; private set;}
        public ChartTitleRenderer Title { get; set; }
        public double Min { get; set; }
        public double Max { get; set; }
        public double MajorUnit { get; set; }
        public double MinorUnit { get; set; }
        public eTimeUnit? MajorDateUnit { get; set; }
        public eTextOrientation LabelOrientation { get; set; }
        public bool IsDateAutoAxis { get; set; }
        public bool IsDateScale
        {
            get;
            private set;
        } = false;
        public bool IsCount { get; set; }

        public override void AppendRenderItems(List<Transform> renderItems)
        {
            var plotareaGroup = ChartRenderer.Plotarea.Group;

            if (Line != null)
            {

                var axisGroup = new GroupRenderItem(ChartRenderer.Plotarea.Rectangle);
                axisGroup.Name = $"Axis_{Axis.Index}";
                axisGroup.Top = Line.Y1;
                axisGroup.Left = Line.X1;
                if (Axis.IsVertical)
                {
                    Line.X1 = Line.X2 = 0;
                }
                else
                {
                    Line.Y1 = Line.Y2 = 0;
                }

                axisGroup.ChildObjects.Add(Line);
                var TickMarkGroup = new GroupRenderItem(axisGroup);
                TickMarkGroup.Name = $"Axis_{Axis.Index}_TickMarkGroup";

                AddSubGroupingOfRenderItems("MinorTickMarkPositions", MinorTickMarkPositions, TickMarkGroup);
                AddSubGroupingOfRenderItems("MajorTickMarkPositions", MajorTickMarkPositions, TickMarkGroup);

                if (MinorTickMarkPositions != null || MajorTickMarkPositions != null)
                {
                    axisGroup.ChildObjects.Add(TickMarkGroup);
                }

                //renderItems.Add(axisGroup);
                plotareaGroup.AddChildItem(axisGroup);
            }
            Title?.AppendRenderItems(renderItems);

            //The axis text boxes is rendered later as they have a higher Z-order.
        }


        internal void AddTickmarksAndValues(List<RenderItem> DefItems)
        {
            var IsCatAxDataTable = (Axis.ActualAxisPosition == eAxisPosition.Bottom && ChartRenderer.HasDataTable && ChartRenderer.DataTable.IsHeaderBottomAxis); //Do not print axis textboxes if the data table header is the text boxes.
            if (Axis.Deleted == false)
            {
                if((Axis.LogBase ?? 1) > 1)
                {
                    ShrinkList(AxisValues, AxisMaxItems);
                }

                if (Axis.MajorTickMark != eAxisTickMark.None && !IsCatAxDataTable)
                {
                    MajorTickMarkPositions = AddTickmarks(MajorUnit, MajorDateUnit, double.NaN, 4D.PixelToPoint(), Axis.MajorTickMark);
                }

                if (Axis.MinorTickMark != eAxisTickMark.None && MinorUnit < MajorUnit && (Axis.LogBase.HasValue==false || Axis.LogBase > 2) && !IsCatAxDataTable)
                {
                    MinorTickMarkPositions = AddTickmarks(MinorUnit, MajorDateUnit, MajorUnit, 2D.PixelToPoint(), Axis.MinorTickMark);
                }
                else
                {
                    MinorTickMarkPositions = null;
                }
            }

            if(Axis.HasMajorGridlines)
            {
                ChartRenderer.Plotarea.MajorGridlinePositions = AddGridlines(MajorUnit, double.NaN, Axis.MajorGridlines, Chart.StyleManager.Style?.GridlineMajor);
            }

            if ((Axis.HasMinorGridlines))
            {
                ChartRenderer.Plotarea.MinorGridlinePositions = AddGridlines(MinorUnit, MajorUnit, Axis.MinorGridlines, Chart.StyleManager.Style?.GridlineMinor);
            }

            if (AxisValues != null && 
                AxisValues.Count > 0 && 
                Axis.Deleted==false && 
                Axis.LabelPosition != eTickLabelPosition.None &&
                !IsCatAxDataTable) //Do not print axis textboxes if the data table header is the text boxes.
            {
                Textboxes = new ChartAxisTextBoxes(this)
                {
                    AxisName = $"Axis_{Axis.Index}_Textboxes",
                    TextBoxes = GetAxisValueTextBoxes(out double widest, out double highest)
                };

                //Adjust position of the axis textboxes so they adjust to the width or height
                if (AdjustPosBySize)
                {
                    if (Axis.IsVertical)
                    {     
                        if(ChartRenderer.VerticalAxis==this || ChartRenderer.VerticalAxis.Axis.AxisTextboxPosition != Axis.AxisTextboxPosition)
                        {
                            Rectangle.Left -= Rectangle.Width;
                        }
                        else
                        {

                        }
                    }
                    else
                    {
                        if (ChartRenderer.HorizontalAxis == this || ChartRenderer.HorizontalAxis.Axis.AxisTextboxPosition != Axis.AxisTextboxPosition)
                        {
                            Rectangle.Top -= Rectangle.Height;
                        }
                        else
                        {
                            if(Axis.AxisTextboxPosition == eActualAxisPosition.Top)
                            {
                                Rectangle.Top = ChartRenderer.HorizontalAxis.Rectangle.Top - TopMargin - Rectangle.Height;
                            }
                            else
                            {
                                Rectangle.Top = ChartRenderer.HorizontalAxis.Rectangle.Bottom + TopMargin;
                            }
                        }
                    }
                }
                
                if (Axis.IsVertical && ChartRenderer.VerticalAxis != this && ChartRenderer.VerticalAxis.Axis.AxisTextboxPosition == Axis.AxisTextboxPosition)
                {
                    if (Axis.AxisTextboxPosition == eActualAxisPosition.Right)
                    {
                        Rectangle.Left = ChartRenderer.VerticalAxis.Rectangle.Right + LeftMargin;
                    }
                    else
                    {
                        Rectangle.Left = ChartRenderer.VerticalAxis.Rectangle.Left - Rectangle.Width - LeftMargin;
                    }
                }
                else if(Axis.IsVertical==false && ChartRenderer.HorizontalAxis != this && ChartRenderer.HorizontalAxis.Axis.AxisTextboxPosition == Axis.AxisTextboxPosition)
                {
                    if (Axis.AxisTextboxPosition == eActualAxisPosition.Bottom)
                    {
                        Rectangle.Top = ChartRenderer.HorizontalAxis.Rectangle.Bottom + TopMargin;
                    }
                    else
                    {
                        Rectangle.Top = ChartRenderer.HorizontalAxis.Rectangle.Top - Rectangle.Height - TopMargin;
                    }
                }

            }
        }

        private List<DrawingTextBox> GetAxisValueTextBoxes(out double widest, out double highest)
        {
            widest = highest = 0;
            var ret = new List<DrawingTextBox>();
            if (Axis.LabelPosition == eTickLabelPosition.None) return ret;

            //All labels share the axis font. The shaper is obtained through the render context
            //so the family is substituted for the render target, matching the font used in the output.
            var mf = Axis.Font.GetMeasureFont(DefaultFonts2007Style.Axis);
            var shaper = ChartRenderer.RenderContext.GetMeasurementShaper(new FontFormatBase(mf));
            var fontSize = mf.Size;
            var lineHeight = shaper.GetLineHeightInPoints(fontSize);

            var axisStyle = GetAxisStyleEntry();
            double maxWidth, maxHeight;
            if (Axis.AxisPosition == eAxisPosition.Left || Axis.AxisPosition == eAxisPosition.Right)
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
                        maxWidth = ((Rectangle.Width + Rectangle.Height) / COS45) / AxisValues.Count;
                        maxHeight = ChartRenderer.ChartArea.Rectangle.Height / 3; //TODO: Check this value.
                        break;
                    case eTextOrientation.Horizontal:
                        maxWidth = Rectangle.Width / AxisValues.Count;
                        maxHeight = ChartRenderer.ChartArea.Rectangle.Height / 3; //TODO: Check this value.
                        break;
                    default: // custom
                        
                        var radRot = MathHelper.Radians(Axis.TextBody.RotationAdjustedForTextVertical);
                        var sin = Math.Abs(Math.Sin(radRot));
                        var cos = Math.Abs(Math.Cos(radRot));
                        var width = Rectangle.Width;
                        //maxWidth = 200;
                        maxWidth = Rectangle.Width / AxisValues.Count;
                        //maxWidth = ((Rectangle.Width + Rectangle.Height) / cos);
                        //maxWidth = Rectangle.Width / ((Rectangle.Width * Math.Abs(Math.Sin(radRot)) + mf.Size * Math.Abs(Math.Cos(radRot))));
                        //Rectangle.Width = Rectangle.Width
                        maxHeight = ChartRenderer.ChartArea.Rectangle.Height / 3; //TODO: Check this value.
                        break;
                }
            }
            int labelSkip = Axis.AxisType == eAxisType.Cat ? (int)(Axis.MajorUnit??1) : 1;
            for (var i = 0; i < AxisValues.Count; i++)
            {
                if (i % labelSkip != 0) continue;
                var v = Values[i];
                var t = AxisValues[i];

                //An empty label has no size, as with ITextMeasurer.
                float width = string.IsNullOrEmpty(t) ? 0f : shaper.Shape(t).GetWidthInPoints(fontSize);
                float height = string.IsNullOrEmpty(t) ? 0f : lineHeight;

                var ticMarkX = GetAxisItemLeft(i);
                double majorTickMarkWidth;
                if(i + 1 < Values.Count)
                {
                    var nextTicMarkX = GetAxisItemLeft(i + 1);
                    majorTickMarkWidth = nextTicMarkX - ticMarkX;
                }
                else
                {
                    if(i > 0)
                    {
                        var prevTickMarkX = GetAxisItemLeft(i - 1);
                        majorTickMarkWidth = ticMarkX - prevTickMarkX;
                    }
                    else
                    {
                        //There is only one axisValue as we are somehow both the first and last value
                        majorTickMarkWidth = 0;
                    }
                }


                var ticMarkY = GetAxisItemTop(i, width, height);
                double x, y;
                if (LabelOrientation == eTextOrientation.Horizontal)
                {
                    if (Axis.AxisType == eAxisType.Cat || Axis.AxisType == eAxisType.Date)
                    {
                        x = ticMarkX;
                        y = ticMarkY;
                    }
                    else
                    {
                        if (Axis.IsVertical)
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
                            x = ticMarkX - width / 2;
                            y = ticMarkY;
                        }
                    }
                }
                else
                {
                    double rot;
                    if(LabelOrientation==eTextOrientation.Custom)
                    {                        
                        rot = Axis.TextBody?.RotationAdjustedForTextVertical??0D;
                    }
                    else
                    {
                        rot = LabelOrientation == eTextOrientation.Diagonal ? -45 : -90;
                    }
                    double cos = Math.Cos(MathHelper.Radians(rot));
                    double sin = Math.Sin(MathHelper.Radians(rot));

                    if (LabelOrientation == eTextOrientation.Diagonal || (LabelOrientation == eTextOrientation.Custom && rot != 90))
                    {
                        if(width < majorTickMarkWidth)
                        {
                            x = ticMarkX;
                        }
                        else
                        {
                            x = ticMarkX - ((height / 2) * cos);
                        }

                        var midPointText = height / 2;
                        var halfHeightAfterRot = Math.Abs(((width / 2) * sin));
                        var heightPastY = halfHeightAfterRot - midPointText;

                        y = ticMarkY - heightPastY;
                        ////x = ticMarkX - (height / 2) * cos;
                        //if (Axis.ActualAxisPosition == eAxisPosition.Bottom)
                        //{
                        //    y = ticMarkY + TopMargin - Math.Abs(((width/2) * sin));
                        //}
                        //else //Top
                        //{
                        //    y = ticMarkY - BottomMargin - Math.Abs(((width / 2) * sin));
                        //}
                    }
                    else if((rot == 90 || rot == -90))
                    {
                        x = ticMarkX;
                        //Width becomes height post-rotation
                        //We are currently pre-rotation.
                        y = ticMarkY - (width/2);
                    }
                    else 
                    {
                        x = ticMarkX - (height / 2);
                        if (Axis.ActualAxisPosition == eAxisPosition.Bottom)
                        {
                            y = ticMarkY + BottomMargin + 4;
                        }
                        else //Top
                        {
                            y = ticMarkY - TopMargin - 4;
                        }
                    }
                }

                var tb = new DrawingTextBox(ChartRenderer.RenderContext, Chart, Rectangle, DefaultFonts2007Style.Axis, x, y, width, height, maxWidth, maxHeight);
                if (LabelOrientation == eTextOrientation.Diagonal)
                {
                    tb.Rotation = -45;
                    if (Axis.ActualAxisPosition == eAxisPosition.Bottom)
                    {
                        tb.TextAnchor = eTextAnchor.End;
                    }
                }
                else if (LabelOrientation == eTextOrientation.Vertical)
                {
                    if (Axis.TextBody.VerticalText == eTextVerticalType.EastAsianVertical)
                    {
                        tb.Rotation = 90;
                    }
                    else
                    {
                        tb.Rotation = -90;
                    }
                    if (Axis.ActualAxisPosition == eAxisPosition.Bottom)
                    {
                        tb.TextAnchor = eTextAnchor.End;
                    }
                }
                else if (LabelOrientation == eTextOrientation.Custom)
                {
                    tb.Rotation = Axis.TextBody?.RotationAdjustedForTextVertical ?? 0D;
                    if (Axis.ActualAxisPosition == eAxisPosition.Bottom)
                    {
                        tb.TextAnchor = eTextAnchor.End;
                    }
                }

                var p = Axis.TextBody.Paragraphs.FirstOrDefault();

                if (p != null)
                {
                    if (p.HorizontalAlignment != eTextAlignment.Center && Axis.AxisType != eAxisType.Val && (Axis.AxisPosition == eAxisPosition.Bottom || Axis.AxisPosition == eAxisPosition.Top))
                    {
                        //Horizontal axises are always center aligned visually
                        //Should be broken out as input to ImportParagraph instead of changing the base item
                        p.HorizontalAlignment = eTextAlignment.Center;
                    }
                    tb.ImportParagraph(p, 0, t);
                }
                else
                {
                    tb.AddText(t);
                }

                tb.CenterRotationPoint();

                tb.Rectangle.Style.SetDrawingPropertiesFill(ChartRenderer.Theme, Axis.Fill, axisStyle?.FillReference.Color, UserSpaceSettings.UserSpaceOnUse_Global, DefaultFillColor);

                if (widest < tb.Width)
                {
                    widest = tb.Width;
                }
                if(highest < tb.Height)
                {
                    highest = tb.Height;
                }
                ret.Add(tb);
            }

            if (Axis.IsVertical)
            {
                //If the axis is vertical, we need to adjust the left position of the textboxes to align them to the right and not have them overlap with the axis line.
                if (Axis.AxisTextboxPosition == eActualAxisPosition.Left)
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
            else if (LabelOrientation == eTextOrientation.Horizontal && IsCatAx()) //Only apples when labels are horizontally aligned
            {
                //Align the axis labels according to the label alignment setting. This is only relevant for horizontal axis, vertical axis are always right aligned.
                var lblAlignment = (Axis as ExcelChartAxisStandard)?.LabelAlignment ?? OfficeOpenXml.eAxisLabelAlignment.Center;
                double majorWidth;
                if (IsDateAutoAxis || IsDateScale)
                {
                    var min = ConvertUtil.GetValueDouble(Values[0]);
                    var max = ConvertUtil.GetValueDouble(Values.Last());
                    var minUnit = (max - min+1) / MinorUnit;
                    majorWidth = Rectangle.Width / minUnit;
                }
                else
                {
                    majorWidth = Rectangle.Width / AxisValues.Count;
                }
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
            else if (LabelOrientation == eTextOrientation.Diagonal)
            {
                if (!(Axis.CrossingAxis == null || Axis.CrossingAxis.CrossBetween == eCrossBetween.MidCat))
                {
                    double majorWidth;
                    if (IsDateAutoAxis || IsDateScale)
                    {
                        var min = ConvertUtil.GetValueDouble(Values[0]);
                        var max = ConvertUtil.GetValueDouble(Values.Last());
                        var minUnit = (max - min) / MinorUnit;
                        majorWidth = Rectangle.Width / minUnit;
                    }
                    else
                    {
                        majorWidth = Rectangle.Width / AxisValues.Count;
                    }
                    foreach (var tb in ret)
                    {
                        tb.Left += majorWidth / 2;
                    }
                }
            }
            else if(LabelOrientation==eTextOrientation.Horizontal && IsCatAx())
            {
                double majorWidth;
                var lblAlignment = (Axis as ExcelChartAxisStandard)?.LabelAlignment ?? OfficeOpenXml.eAxisLabelAlignment.Center;
                if (IsDateAutoAxis || IsDateScale)
                {
                    var min = ConvertUtil.GetValueDouble(Values[0]);
                    var max = ConvertUtil.GetValueDouble(Values.Last());
                    var minUnit = (max - min) / MinorUnit;
                    majorWidth = Rectangle.Width / minUnit;
                }
                else
                {
                    majorWidth = Rectangle.Width / AxisValues.Count;
                }
                foreach (var tb in ret)
                {
                    if(tb.Width > majorWidth) //If the box is smaller than the text box with, center the label.
                    {
                        tb.Left +=  -tb.Width / 2;
                    }
                    else
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

            return ret;
        }

        private bool IsCatAx()
        {
            return Axis.AxisType == eAxisType.Cat || (Axis.AxisType == eAxisType.Date && IsDateScale==false);
        }

        private double GetAxisItemLeft(int i)
        {
            if (Axis.IsVertical)
            {
                return 0;
            }
            else
            {
                if (IsCatAx() && IsDateAutoAxis==false) //A text axis
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
                    var majorTickStartingPosition = majorWidth * i;
                    return majorTickStartingPosition;
                }
                else
                {
                    var min = ConvertUtil.GetValueDouble(Values[0]);
                    var max = ConvertUtil.GetValueDouble(Values.Last());
                    var v = ConvertUtil.GetValueDouble(Values[i]);
                    double majorWidth;
                    if (IsDateAutoAxis || IsDateScale)
                    {
                        majorWidth = Rectangle.Width * (v - Min) / (Max - Min + 1);
                    }
                    else
                    {
                        majorWidth = Rectangle.Width * (v - Min) / (Max - Min);
                    }
                    
                    return majorWidth;
                }
                //}
            }
        }

        private double GetAxisItemTop(int i, double width, double height)
        {
            if (Axis.ActualAxisPosition == eAxisPosition.Top)
            {
                switch (LabelOrientation)
                {
                    case eTextOrientation.Vertical:
                    case eTextOrientation.Diagonal:
                        return Rectangle.Height;
                    default:
                        return Rectangle.Height - height - TopMargin;
                }
            }
            else if (Axis.ActualAxisPosition == eAxisPosition.Bottom)
            {
                switch(LabelOrientation)
                {
                    case eTextOrientation.Vertical:
                        if (Axis.LabelPosition == eTickLabelPosition.Low)
                        {
                            return Rectangle.Height - width;
                        }
                        else
                        {
                            return 0;
                        }
                    case eTextOrientation.Diagonal:
                        if (Axis.LabelPosition == eTickLabelPosition.Low)
                        {
                            return Rectangle.Height - (width + height)* COS45;
                        }
                        else
                        {
                            return 0;
                        }
                    default:
                        if (Axis.LabelPosition == eTickLabelPosition.Low)
                        {
                            return Rectangle.Height - height;
                        }
                        else if (Axis.LabelPosition == eTickLabelPosition.NextTo)
                        {
                            return BottomMargin;
                        }
                        else //TODO:Add support for hight.
                        {
                            //var top = ChartRenderer.Plotarea.Group.Top + height;
                            //Rectangle.Height += Rectangle.Height - top;
                            //Rectangle.Top = top;
                            return 0;
                        }
                }
            }
            else
            {
                if (Axis.AxisType == eAxisType.Cat || Axis.AxisType == eAxisType.Date)
                {
                    var majorHeight = Rectangle.Height / (AxisValues.Count);
                    return majorHeight * (AxisValues.Count - i) - ((majorHeight / 2) + height / 2);
                }
                else
                {
                    var majorHeight = Rectangle.Height / (AxisValues.Count-1);
                    if(Axis.Orientation==eAxisOrientation.MinMax)
                    {
                        return majorHeight * (AxisValues.Count - i - 1);
                    }
                    else
                    {
                        return majorHeight * i;
                    }
                }
            }

        }

        private List<LineRenderItem> AddTickmarks(double units, eTimeUnit? dateUnit, double parentUnit, double tickMarkWidth, eAxisTickMark type)
        {
            var axisStyle = GetAxisStyleEntry();

            if((Axis.LogBase ?? 1) > 1)
            {
                return AddTickmarksLogarithmic(units, parentUnit, tickMarkWidth, type, axisStyle);
            }
            else
            {
                return AddTickmarksLinear(units, dateUnit, tickMarkWidth, type, axisStyle);
            }
        }

        private List<LineRenderItem> AddTickmarksLogarithmic(double units, double parentUnit, double tickMarkWidth, eAxisTickMark type, ExcelChartStyleEntry axisStyle)
        {
            var tms = new List<LineRenderItem>();

            GetTickmarkWidths(type, tickMarkWidth, out double tickMarkWidthInside, out double tickMarkWidthOutside);
            var diff = Max - Min;
            var logBase = (Axis.LogBase ?? 1);
            var isLogBase = logBase > 1;
            double d = Min;
            double minorUnit = 1;
            var maxPos = Max + units * 1e-12; //To avoid rounding errors when comparing d to maxPos, we add a small epsilon value to the comparison.
            while (d <= maxPos)
            {
                var addPosition = (d - Min);
                double x1, y1, x2, y2;
                switch (Axis.ActualAxisPosition)
                {
                    case eAxisPosition.Left:
                    case eAxisPosition.Right:
                    //    y1 = GetPositionInPlotarea(d);
                    //    y2 = y1;
                    //    x1 = (float)Rectangle.Width - tickMarkWidthOutside;
                    //    x2 = (float)Rectangle.Width + tickMarkWidthInside;
                    //    break;
                    //case eAxisPosition.Right:
                        y1 = GetPositionInPlotarea(d);
                        y2 = y1;
                        x1 = (float)-tickMarkWidthInside;
                        x2 = (float)tickMarkWidthOutside;
                        break;
                    case eAxisPosition.Top:
                    case eAxisPosition.Bottom:
                    //    x1 = GetPositionInPlotarea(d);
                    //    x2 = x1;
                    //    y1 = (float)Rectangle.Height - tickMarkWidthOutside;
                    //    y2 = (float)Rectangle.Height + tickMarkWidthInside;
                    //    break;
                    //case eAxisPosition.Bottom:
                        x1 = GetPositionInPlotarea(d);
                        x2 = x1;
                        y1 = (float)-tickMarkWidthInside;
                        y2 = (float)tickMarkWidthOutside;
                        break;
                    default:
                        throw new InvalidOperationException("Invalid axis position");
                }
                var tm = new LineRenderItem(ChartRenderer.Bounds);
                tm.X1 = x1;
                tm.Y1 = y1;
                tm.X2 = x2;
                tm.Y2 = y2;
                tm.Style.SetDrawingPropertiesBorder(ChartRenderer.Theme, Axis.Border, axisStyle?.BorderReference.Color, true, GetDefaultBorderColor, 0.75);
                if (tm.Style.BorderWidth < 0.75) //Excel seems to have this as minimum width for tick marks, so we enforce it here to make sure they are visible.
                {
                    tm.Style.BorderWidth = 0.75;
                }
                tms.Add(tm);

                if (units == 0) break;
                if (double.IsNaN(parentUnit))
                {
                    d *= units;
                }
                else
                {
                    d += minorUnit;
                    if ((d%parentUnit)==0)
                    {
                        minorUnit *= parentUnit;
                    }
                }
            }
            ShrinkList(tms, AxisMaxItems);
            return tms;
        }

        private List<LineRenderItem> AddTickmarksLinear(double units, eTimeUnit? dateUnit, double tickMarkWidth, eAxisTickMark type, ExcelChartStyleEntry axisStyle)
        {
            var tms = new List<LineRenderItem>();
            GetTickmarkWidths(type, tickMarkWidth, out double tickMarkWidthInside, out double tickMarkWidthOutside);

            double min, max;
            if (Axis.AxisType == eAxisType.Cat && IsDateAutoAxis == false)
            {
                if (AxisValues != null)
                {
                    min = 1;
                    if (Axis.CrossingAxis == null || Axis.CrossingAxis.CrossBetween == eCrossBetween.Between)
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
                    min = 0;
                    max = 0;
                }
            }
            else
            {
                min = Min;
                max = Max;
            }

            var diff = IsCatAx() ? max - min + 1 : max - min;
            var maxPos = IsCatAx() ? max + 1 : max;
            double d = min;
            maxPos += units * 1e-12; //To avoid rounding errors when comparing d to maxPos, we add a small epsilon value to the comparison.
            int skipTickmark = (Axis.AxisType == eAxisType.Cat ? (int)(Axis.MinorUnit ?? 1): 1);
            while (d <= maxPos)
            {
                var addPosition = (d - min);
                if (addPosition % skipTickmark == 0)
                {
                    double x1, y1, x2, y2;
                    switch (Axis.ActualAxisPosition)
                    {
                        case eAxisPosition.Left:
                            y1 = (float)(Rectangle.Height - (addPosition / diff * Rectangle.Height));
                            y2 = y1;
                            x1 = (float)-tickMarkWidthOutside;
                            x2 = (float)+tickMarkWidthInside;
                            break;
                        case eAxisPosition.Right:
                            y1 = (float)(Rectangle.Height - (addPosition / diff * Rectangle.Height));
                            y2 = y1;
                            x1 = (float)-tickMarkWidthInside;
                            x2 = (float)tickMarkWidthOutside;
                            break;
                        case eAxisPosition.Top:
                            x1 = (float)((addPosition / diff * Rectangle.Width));
                            x2 = x1;
                            y1 = (float)-tickMarkWidthOutside;
                            y2 = (float)+tickMarkWidthInside;
                            break;
                        case eAxisPosition.Bottom:
                            x1 = (float)((addPosition / diff * Rectangle.Width));
                            x2 = x1;
                            y1 = (float)-tickMarkWidthInside;
                            y2 = (float)tickMarkWidthOutside;
                            break;
                        default:
                            throw new InvalidOperationException("Invalid axis position");
                    }
                    var tm = new LineRenderItem(ChartRenderer.Bounds);
                    tm.X1 = x1;
                    tm.Y1 = y1;
                    tm.X2 = x2;
                    tm.Y2 = y2;
                    tm.Style.SetDrawingPropertiesBorder(ChartRenderer.Theme, Axis.Border, axisStyle?.BorderReference.Color, true, GetDefaultBorderColor, 0.75);
                    if (tm.Style.BorderWidth < 0.75) //Excel seems to have this as minimum width for tick marks, so we enforce it here to make sure they are visible.
                    {
                        tm.Style.BorderWidth = 0.75;
                    }
                    tms.Add(tm);
                }
                if (units == 0) break;
                switch (dateUnit)
                {
                    case eTimeUnit.Years:
                        d = DateTime.FromOADate(d).AddYears((int)units).ToOADate();
                        break;
                    case eTimeUnit.Months:
                        if (units >= 1D)
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

        private void GetTickmarkWidths(eAxisTickMark type, double tickMarkWidth, out double tickMarkWidthInside, out double tickMarkWidthOutside)
        {
            if (type == eAxisTickMark.In || type == eAxisTickMark.Cross)
            {
                tickMarkWidthInside = tickMarkWidth;
            }
            else
            {
                tickMarkWidthInside = 0;
            }

            if (type == eAxisTickMark.Out || type == eAxisTickMark.Cross)
            {
                tickMarkWidthOutside = tickMarkWidth;
            }
            else
            {
                tickMarkWidthOutside = 0;
            }

        }

        private bool IsMinorDateUnit(eTimeUnit dateUnit, double parentUnit, double d)
        {
            switch(dateUnit)
            {
                case eTimeUnit.Days:
                    return d % parentUnit != 0;
                case eTimeUnit.Months:
                    var minDt = DateTime.FromOADate(Min);
                    var dt = DateTime.FromOADate(d);
                    return minDt.Month % parentUnit != dt.Month % parentUnit;
                case eTimeUnit.Years:
                    minDt = DateTime.FromOADate(Min);
                    dt = DateTime.FromOADate(d);
                    return minDt.Year % parentUnit != dt.Year % parentUnit;
                default:
                    throw new InvalidOperationException("Invalid date unit");
            }
        }

        private List<RenderItem> AddGridlines(double units, double parentUnit, ExcelDrawingBorder lineItem, ExcelChartStyleEntry styleEntry)
        {
            var axisStyle = GetAxisStyleEntry();

            var points = GetGridLinePositions(units, parentUnit);

            float x1, y1, x2, y2;

            string id = "";
            var pa = ChartRenderer.Plotarea;
            switch (Axis.AxisPosition)
            {
                case eAxisPosition.Left:
                case eAxisPosition.Right:
                    id = $"x{(double.IsNaN(parentUnit) ? "Major" : "Minor")}GridLine";
                    y1 = (float)points.Last().Top;
                    y2 = y1;
                    x1 = (float)0;
                    x2 = (float)pa.Rectangle.Width;
                    points.Reverse(); //Reverse, so the for loop below gets the items in the correct order.
                    break;
                case eAxisPosition.Top:
                case eAxisPosition.Bottom:
                    id = $"y{(double.IsNaN(parentUnit) ? "Major" : "Minor")}GridLine";
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
            tm.Style.SetDrawingPropertiesBorder(ChartRenderer.Theme, lineItem, styleEntry?.BorderReference.Color, true, GetDefaultBorderColor, 0.75);
            tm.Style.DefId = id;

            var tms = new List<RenderItem>();

            tms.Add(tm);

            for (int i = 1; i < points.Count; i++)
            {
                var refItem = new UseReferenceRenderItem(ChartRenderer.Bounds, "#" + id);
                if (Axis.IsVertical)
                {
                    refItem.X = 0f;
                    refItem.Y = points[i].Top;
                }
                else
                {
                    refItem.X = points[i].Left;
                    refItem.Y = 0f;
                }
                tms.Add(refItem);
            }


            return tms;
        }

        private List<TranformPoint> GetGridLinePositions(double units, double parentUnit)
        {
            if ((Axis.LogBase ?? 1) > 1)
            {
                var list = GetGridLinePositionsLogarithmic(units, parentUnit);
                if (double.IsNaN(parentUnit))
                {
                    ShrinkList(list, AxisMaxItems);
                }
                return list;
            }
            return GetGridLinePositionsLinear(units);
        }

        private List<TranformPoint> GetGridLinePositionsLogarithmic(double units, double parentUnit)
        {
            var pa = ChartRenderer.Plotarea;
            var points = new List<EPPlus.Graphics.TranformPoint>();
            var group = ChartRenderer.Plotarea.Group;
            var minorUnit = 1D;
            var maxPos = Max + units * 1e-12; //To avoid rounding errors when comparing d to maxPos, we add a small epsilon value to the comparison.
            var d = Min;
            while (d <= maxPos)
            {
                switch (Axis.AxisPosition)
                {
                    case eAxisPosition.Left:
                    case eAxisPosition.Right:
                        var y = GetPositionInPlotarea(d);
                        points.Add(new TranformPoint(0f, y));
                        break;
                    case eAxisPosition.Top:
                    case eAxisPosition.Bottom:
                        var x = GetPositionInPlotarea(d);
                        points.Add(new TranformPoint(x, 0f));
                        break;
                    default:
                        throw new InvalidOperationException("Invalid axis position.");
                }
                if (units == 0) break;
                if(double.IsNaN(parentUnit))
                {
                    d *= units;
                }
                else
                {
                    if (double.IsNaN(parentUnit))
                    {
                        d *= units;
                    }
                    else
                    {
                        d += minorUnit;
                        if ((d % (minorUnit * parentUnit)) == 0)
                        {
                            minorUnit *= parentUnit;
                        }
                    }
                }
            }
            return points;
        }

        private List<TranformPoint> GetGridLinePositionsLinear(double units)
        {
            double min;
            if (IsCount)
            {
                min = 0;
            }
            else
            {
                min = Min;
            }
            var pa = ChartRenderer.Plotarea;
            var diff = Max - min;

            var points = new List<EPPlus.Graphics.TranformPoint>();
            var group = ChartRenderer.Plotarea.Group;
            for (double d = min; d <= Max + units * 1e-12; d += units)
                {
                switch (Axis.AxisPosition)
                {
                    case eAxisPosition.Left:
                    case eAxisPosition.Right:
                        points.Add(new EPPlus.Graphics.TranformPoint(0f, (float)(pa.Rectangle.Height - ((d - min) / diff * pa.Rectangle.Height))));
                        break;
                    case eAxisPosition.Top:
                    case eAxisPosition.Bottom:
                        var xValue = (float)(((d - min) / diff * pa.Rectangle.Width));
                        points.Add(new EPPlus.Graphics.TranformPoint(xValue, 0f));
                        break;
                    default:
                        throw new InvalidOperationException("Invalid axis position.");
                }
                if (units == 0) break;
            }
            return points;
        }

        private ExcelChartStyleEntry GetAxisStyleEntry()
        {
            ExcelChartStyleEntry axisStyle;
            switch (Axis.AxisType)
            {
                case eAxisType.Cat:
                case eAxisType.Date:
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
        private double GetPositionInPlotareaLog(double val, bool startValue = false)
        {
            var diff = (Max - Min) + 1;
            double logMax = Math.Log(diff, Axis.LogBase.Value);
            double fraction = Math.Log(val, Axis.LogBase.Value);

            if (Axis.AxisPosition == eAxisPosition.Left || Axis.AxisPosition == eAxisPosition.Right)
            {
                return ChartRenderer.Plotarea.Rectangle.Height-(fraction / logMax * ChartRenderer.Plotarea.Rectangle.Height);
            }
            else
            {
                return fraction / logMax * ChartRenderer.Plotarea.Rectangle.Width;
            }                    
        }
        internal double GetPositionInPlotarea(double val, bool startValue=false, double slotSize=1)
        {
            double pos;
            if (Axis.LogBase.HasValue)
            {
                pos = GetPositionInPlotareaLog(val, startValue);
            }
            else
            {
                pos = GetPositionInPlotAreaLinear(val, startValue);
            }

            if(Axis.Orientation==eAxisOrientation.MaxMin)
            {
                if(Axis.IsVertical)
                {
                    return ChartRenderer.Plotarea.Rectangle.Height-pos;
                }
                else
                {
                    //return ChartRenderer.Plotarea.Rectangle.Width - pos;
                }
            }
            return pos;
        }

        private double GetPositionInPlotAreaLinear(double val, bool startValue)
        {
            if (Axis.AxisPosition == eAxisPosition.Left || Axis.AxisPosition == eAxisPosition.Right)
            {
                if (AutoAxisType == eAxisType.Cat && IsDateAutoAxis == false)
                {
                    var majorHeight = ChartRenderer.Plotarea.Rectangle.Height / Max;
                    if (startValue)
                    {
                        return majorHeight * val;
                    }
                    else
                    {
                        return majorHeight * val + (majorHeight / 2);
                    }
                }
                else if ((AutoAxisType == eAxisType.Date) && IsDateScale == false)
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
                if (Axis.AxisType == eAxisType.Cat && IsDateAutoAxis == false)
                {
                    var majorWidth  = ChartRenderer.Plotarea.Rectangle.Width / Max;
                    if (startValue)
                    {
                        return majorWidth * val;
                    }
                    else
                    {
                        if (Axis.CrossingAxis == null || Axis.CrossingAxis.CrossBetween == eCrossBetween.Between)
                        {
                            return majorWidth * (val - 1) + (majorWidth / 2);
                        }
                        else
                        {
                            majorWidth = ChartRenderer.Plotarea.Rectangle.Width / (Max - 1);
                            return majorWidth * (val - 1);
                        }
                    }
                }
                else if ((AutoAxisType == eAxisType.Date) && IsDateScale == false)
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
            var values = ax.GetAxisValues(out bool isCount, out bool isNumeric, out bool isDate);
            IsCount = isCount;
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

            //All axis labels share the axis font. The shaper is obtained through the render context
            //so the family is substituted for the render target, matching the font used in the output.
            var mf = ax.Font.GetMeasureFont(DefaultFonts2007Style.Axis);
            var shaper = ChartRenderer.RenderContext.GetMeasurementShaper(new FontFormatBase(mf));
            var fontSize = mf.Size;

            if (AutoAxisType == eAxisType.Cat && isCount == false && isDate == false)
            {
                AxisScale res;
                if (ax.IsVertical)
                {
                    res = CategoryAxisScaleCalculator.CalculateVerticalAxisByHeight(ref values, shaper, fontSize, options);
                }
                else
                {
                    res = CategoryAxisScaleCalculator.CalculateHorizontalAxisByWidth(ref values, shaper, fontSize, options);
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
                double d;
                object ov;
                if (v is object[] o)
                {
                    ov = o[3];
                }
                else
                {
                    ov = v;
                }
                d = ConvertUtil.GetValueDouble(ov, false, true);
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
            if (isCount)
            {
                for (int i = 1; i <= max; i++)
                {
                    l.Add(i);
                }
                var res = CategoryAxisScaleCalculator.CalculateHorizontalAxisByWidth(ref l, shaper, fontSize, options);

                min = res.Min;
                max = res.Max;
                majorUnit = res.MajorInterval;
                dateUnit = null;
                orientation = res.TextOrientation;

                return l.ToList();
            }
            if (ax.AxisType == eAxisType.Val)
            {
                AdjustminMaxFromChartObjects(ax, ref min, ref max);
            }
            if (isDate || ax.IsDate)
            {
                AxisScale res;
                if (ax.IsVertical)
                {
                    res = DateAxisScaleCalculator.CalculateByWidthHeight(options.ChartSize.Height, min ?? 0D, max ?? 0D, shaper, fontSize, options);
                }
                else
                {
                    if (ax.AxisType == eAxisType.Val)
                    {
                        res = DateAxisScaleCalculator.Calculate(min ?? 0D, max ?? 0D, options);
                    }
                    else
                    {
                        res = DateAxisScaleCalculator.CalculateByWidthAllowDiagonal(values, min ?? 0D, max ?? 0D, shaper, fontSize, options);
                    }
                }

                orientation = res.TextOrientation;
                dateUnit = res.MajorDateUnit;
                majorUnit = res.MajorInterval;
                var dt = DateTime.FromOADate(res.Min);
                var maxDt = DateTime.FromOADate(res.Max);
                IsDateAutoAxis = true;
                IsDateScale = (dateUnit != eTimeUnit.Days || majorUnit > 1) && values.Count > 31;

                while (dt <= maxDt)
                {
                    l.Add(dt);
                    switch (res.MajorDateUnit ?? eTimeUnit.Days)
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
                if (ax.LogBase > 1)
                {
                    for (var v = res.Min; Math.Round(v, 12) <= Math.Round(res.Max, 12); v *= res.MajorInterval)
                    {
                        l.Add(v);
                        if (res.MajorInterval == 0) break;
                    }
                }
                else
                {
                    for (var v = res.Min; Math.Round(v, 12) <= Math.Round(res.Max, 12); v += res.MajorInterval)
                    {
                        l.Add(v);
                        if (res.MajorInterval == 0) break;
                    }
                }


                min = res.Min;
                max = res.Max;
                majorUnit = res.MajorInterval;
                dateUnit = null;
                orientation = options.Axis.TextBody.RotationAdjustedForTextVertical !=0 ? eTextOrientation.Custom : eTextOrientation.Horizontal;
            }

            return l;
        }

        private void ShrinkList<T>(List<T> l, int max)
        {
            while(l.Count > max)
            { 
                var c = l.Count-1;
                for(int i=1;i<c--;i+=1)
                {
                    l.RemoveAt(i);
                }
            }
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

        internal override Color? GetDefaultFillColor()
        {
            return GetDefaultFillColorForElement(ChartRenderer.Theme, ChartElement.Axis, (int)Chart.Style);
        }

        internal override Color? GetDefaultBorderColor()
        {
            return GetDefaultBorderColorForElement(ChartRenderer.Theme, ChartElement.Axis, (int)Chart.Style);
        }
        internal override Color? DefaultBorderColor => GetDefaultBorderColor();



        /// <summary>
        /// If the axis is left or top, the position of the axis is adjusted by the size of the axis. 
        /// </summary>
        internal bool AdjustPosBySize { get; set; } = false;

        internal double GetAdjustedCrossesAt()
        {
            var crossAx = CrossingAxis;
            if(Axis.CrossingAxis !=null && Axis.CrossingAxis.CrossesAt.HasValue)
            {
                var ca = Axis.CrossingAxis.CrossesAt.Value;
                if (ca < crossAx.Min)
                {
                    return crossAx.Min;
                }
                else if(ca > crossAx.Max)
                {
                    return crossAx.Max;
                }
                else
                {
                    if(IsCount)
                    {
                        return ca-1;
                    }
                    return ca;
                }   
            }
            else
            {
                switch (Axis.Crosses)
                {
                    case eCrosses.Min:
                        return crossAx.Min;
                    case eCrosses.Max:
                        return crossAx.Max;
                    default:
                        return Math.Min(Math.Max(0,crossAx.Min), crossAx.Max);
                }
            }
        }
        internal ChartAxisRenderer CrossingAxis
        {
            get
            {
                if (ChartRenderer.HorizontalAxis == this)
                {
                    return ChartRenderer.VerticalAxis;
                }
                else if (ChartRenderer.VerticalAxis == this)
                {
                    return ChartRenderer.HorizontalAxis;
                }
                if (ChartRenderer.SecondHorizontalAxis == this)
                {
                    return ChartRenderer.SecondVerticalAxis;
                }
                else 
                {
                    return ChartRenderer.SecondHorizontalAxis;
                }
            }
        }
        internal double GetCrossesAt()
        {
            if (Axis.CrossesAt.HasValue)
            {
                var ca = Axis.CrossesAt.Value;
                if (ca < Min)
                {
                    return Min;
                }
                else if (ca > Max)
                {
                    return Max;
                }
                else
                {
                    if (IsCount)
                    {
                        return ca - 1;
                    }
                    return ca;
                }
            }
            else
            {
                switch (Axis.Crosses)
                {
                    case eCrosses.Min:
                        return Min;
                    case eCrosses.Max:
                        return Max;
                    default:
                        return Math.Min(Math.Max(0, Min), Max);
                }
            }
        }

        internal eTimeUnit GetBaseUnit()
        {
            if (Axis.BaseTimeUnit.HasValue) return Axis.BaseTimeUnit.Value;
            DateTime? prevDate=null;
            bool isMonth = false, isYear = false;
            int min=int.MaxValue, max=0;
            var nullCount = 0;
            foreach (var v in Axis.GetAxisValues(out _, out _, out _))
            {
                DateTime? d;
                if (v is Array[] a)
                {
                    d = ConvertUtil.GetValueDate(a[3]);
                }
                else
                {
                    d= ConvertUtil.GetValueDate(v);
                }
                if (d == null)
                {
                    nullCount++;
                    continue;
                }
                if (prevDate.HasValue)
                {
                    var dt = d.Value;
                    var pdt = prevDate.Value;
                    if (dt.Year == pdt.Year && dt.Month == pdt.Month)
                    {
                        return eTimeUnit.Days;
                    }
                    else if(dt.Month == pdt.AddMonths(nullCount + 1).Month && dt.Year == pdt.AddMonths(nullCount + 1).Year)
                    {
                        min = Math.Min(min, d.Value.Day);
                        max = Math.Max(max, d.Value.Day);
                        if (max - min > 4) return eTimeUnit.Days;
                        isMonth = true;
                    }
                    else if (d.Value.Year == prevDate.Value.AddYears((nullCount + 1)).Year && isMonth==false)
                    {
                        isYear = true;
                    }
                    else
                    {
                        isMonth = true;
                    }
                }
                nullCount = 0;
                prevDate = d.Value;
            }
            if(isYear && isMonth==false)
            {
                return eTimeUnit.Years;
            }
            if(isMonth)
            {
                return eTimeUnit.Months;
            }
            return eTimeUnit.Days;
        }
    }
}