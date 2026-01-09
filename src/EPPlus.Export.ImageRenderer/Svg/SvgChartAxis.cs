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
using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartAxis : SvgChartObject, IDrawingChartAxis
    {
        internal SvgChartAxis(SvgChart sc, ExcelChartAxisStandard ax) : base(sc.Chart)
        {

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

            if (ax.Deleted == false)
            {
                Values = GetAxisValue(ax, Rectangle, out double? min, out double? max, out double? majorUnit);
                AxisValues = GetAxisDisplayValues(ax, Values, min, max, majorUnit);
                
                Min = min??0D;
                Max = max.Value;
                MajorUnit = majorUnit??1;

                if (ax.Layout.HasLayout)
                {
                    Rectangle = GetRectFromManualLayout(sc, ax.Layout);
                }
                else
                {
                    Rectangle = new SvgRenderRectItem(sc.Chart);
                    if (ax.AxisPosition == eAxisPosition.Left || ax.AxisPosition == eAxisPosition.Right)
                    {
                        Rectangle.Width = GetTextWidest(sc, ax);
                        Rectangle.Left = Title == null || ax.AxisPosition == eAxisPosition.Right ? 8D : Title.Rectangle.Right;
                    }
                    else
                    {
                        Rectangle.Height = GetTextHeight(sc, ax);
                        Rectangle.Top = Title == null || ax.AxisPosition == eAxisPosition.Top ? sc.ChartArea.Height - 8 - Rectangle.Height : Title.Rectangle.Top - Rectangle.Height - 8;
                    }
                }

                Rectangle.FillColor = "none";

                Line = new SvgRenderLineItem(sc.Chart);
                Line.SetDrawingPropertiesBorder(ax.Border, sc.Chart.StyleManager.Style.Title.BorderReference.Color, ax.Border.Fill.Style != eFillStyle.NoFill, 0.75);
            }
        }

        private List<SvgRenderLineItem> GetMajorAxisItems(double? min, double? max, double? majorUnit)
        {
            var mtms = new List<SvgRenderLineItem>();
            for (var i = min; i < max; i += majorUnit)
            {
                float x, y;
                switch (Axis.AxisPosition)
                {
                    case eAxisPosition.Left:
                        break;
                    case eAxisPosition.Right:
                        break;
                    case eAxisPosition.Top:
                        break;
                    case eAxisPosition.Bottom:
                        x = Line.X1;
                        break;
                }
            }
            return mtms;
        }

        internal ExcelChartAxis Axis { get; }
        private List<string> GetAxisDisplayValues(ExcelChartAxisStandard ax, List<object> values, double? min, double? max, double? majorUnit)
        {
            var displayValues = new List<string>();
            var nf = new ExcelFormatTranslator(ax.Format, 0);

            foreach(var v in values)
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
        private double GetAxisXPosition(SvgChart sc, ExcelChartAxisStandard ax)
        {
            switch (ax.AxisPosition)
            {
                case eAxisPosition.Left:
                case eAxisPosition.Right:
                    if (sc.Chart.Legend != null && sc.Chart.Legend.Position == eLegendPosition.Top)
                    {
                        return sc.Legend.Rectangle.Bottom + TopMargin;
                    }
                    else
                    {
                        return sc.Title?.Rectangle?.Bottom ?? 0D + TopMargin;
                    }
                case eAxisPosition.Top:
                case eAxisPosition.Bottom:
                    if (sc.Chart.Legend != null && sc.Chart.Legend.Position == eLegendPosition.Bottom)
                    {
                        return sc.Legend.Rectangle.Bottom + BottomMargin;
                    }
                    else
                    {
                        return BottomMargin;
                    }
            }
            return 0;
        }

        private double GetAxisYPosition(SvgChart sc, ExcelChartAxisStandard ax)
        {
            switch (ax.AxisPosition)
            {
                case eAxisPosition.Left:
                case eAxisPosition.Right:
                    if(sc.Chart.Legend != null && sc.Chart.Legend.Position==eLegendPosition.Top)
                    {
                        return sc.Legend.Rectangle.Bottom + TopMargin;
                    }
                    else
                    {
                        return sc.Title?.Rectangle?.Bottom ?? 0D + TopMargin;
                    }
                case eAxisPosition.Top:
                case eAxisPosition.Bottom:
                    if (sc.Chart.Legend != null && sc.Chart.Legend.Position == eLegendPosition.Bottom)
                    {
                        return sc.Legend.Rectangle.Bottom + BottomMargin;
                    }
                    else
                    {
                        return BottomMargin;
                    }
            }
            return 0;
        }

        private double GetAxisYPosition(SvgChart sc)
        {
            throw new NotImplementedException();
        }

        public List<object> Values
        {
            get;
            private set;
        }
        public List<string> AxisValues
        {
            get;
            private set;
        }
        public List<SvgRenderLineItem> MajorAxisPositions { get; private set; }
        public List<SvgRenderLineItem> MinorAxisPositions { get; private set; }
        public SvgChartTitle Title { get; set; }
        public double Min { get; set; }
        public double Max { get; set; }
        public double MajorUnit { get; set; }
        public double MinorUnit { get; set; }
        public override void Render(StringBuilder sb)
        {
            Title?.Render(sb);
            Rectangle?.Render(sb);            
            Line?.Render(sb);
            if (MajorAxisPositions != null)
            {
                foreach (var tm in MajorAxisPositions)
                {
                    tm.Render(sb);
                }
            }
            for(var i=0;i < AxisValues.Count; i++)
            {
                //RenderAxisValue(i);
            }
        }

        private void RenderMajorTickmarks(int i)
        {
            
        }

        internal void AddTickmarks()
        {            
            if (Axis.MajorTickMark != eAxisTickMark.None)
            {
                MajorAxisPositions = AddMajorAxisItems();
            }

            if (Axis.MinorTickMark != eAxisTickMark.None)
            {
                MinorAxisPositions = AddMajorAxisItems();
            }

            if (Axis.CrossBetween == eCrossBetween.MidCat)
            {

            }
            else
            {

            }
        }

        private List<SvgRenderLineItem> AddMajorAxisItems()
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

            var tms=new List<SvgRenderLineItem>();
            for(double d=Min; d<=Max; d+=MajorUnit)
            {
                float x1, y1, x2, y2;
                switch (Axis.AxisPosition)
                {
                    case eAxisPosition.Left:
                        x1 = (float)Rectangle.Right;
                        y1 = (float)(Rectangle.Top + Rectangle.Height - ((d - Min) / (Max - Min) * Rectangle.Height));
                        x2 = x1 + 4;
                        y2 = y1;
                        break;
                    case eAxisPosition.Right:
                        x1 = (float)Rectangle.Left;
                        y1 = (float)(Rectangle.Top + Rectangle.Height - ((d - Min) / (Max - Min) * Rectangle.Height));
                        x2 = x1 - 4;
                        y2 = y1;
                        break;
                    case eAxisPosition.Top:
                        x1 = (float)(Rectangle.Left + ((d - Min) / (Max - Min) * Rectangle.Width));
                        y1 = (float)Rectangle.Bottom;
                        x2 = x1;
                        y2 = y1 - 4;
                        break;
                    case eAxisPosition.Bottom:
                        x1 = (float)(Rectangle.Left + ((d - Min) / (Max - Min) * Rectangle.Width));
                        y1 = (float)Rectangle.Top;
                        x2 = x1;
                        y2 = y1 + 4;
                        break;
                    default:
                        throw new InvalidOperationException("Invalid axis position");
                }
                var tm = new SvgRenderLineItem(Chart);
                tm.X1 = x1;
                tm.Y1 = y1;
                tm.X2 = x2;
                tm.Y2 = y2;
                tm.SetDrawingPropertiesBorder(Axis.Border, axisStyle.BorderReference.Color, true);
                tms.Add(tm);
            }
            return tms;
        }
    }
}