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
        public SvgChartTitle Title { get; set; }
        public override void Render(StringBuilder sb)
        {
            Title?.Render(sb);
            Rectangle?.Render(sb);
            Line?.Render(sb);
        }
    }
}