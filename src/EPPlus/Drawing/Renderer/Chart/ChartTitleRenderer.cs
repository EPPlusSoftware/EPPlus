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
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using EPPlus.DrawingRenderer.RenderItems;

namespace EPPlusImageRenderer.Svg
{
    internal class ChartTitleRenderer : ChartDrawingObject
    {
        ExcelChartTitleStandard _title;
        string _titleText;
        ChartRenderer _svgChart;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sc"></param>
        /// <param name="t"></param>
        /// <param name="defaultText"></param>
        /// <param name="axis">If null, this is the main chart title.</param>
        internal ChartTitleRenderer(ChartRenderer sc, ExcelChartTitleStandard t, string defaultText, SvgChartAxis axis=null) : base(sc)
        {
            _svgChart = sc;

            //These are hard coded margins for the title box.
            LeftMargin = RightMargin = 3; //4px
            TopMargin = BottomMargin = 1.5; //2px
            LeftMargin = RightMargin = 3; //4px
            TopMargin = BottomMargin = 1.5; //2px

            var maxWidth = sc.Bounds.Width * 0.8;
            var maxHeight = sc.Bounds.Height / 2D;
            _title = t;
            if (axis==null)
            {
                _titleText = GetDefaultChartTitleText(sc, t, defaultText);
            }
            else
            {
                if (string.IsNullOrEmpty(t.DisplayedText) == false)
                {
                    _titleText = t.DisplayedText;
                }
                else
                {
                    _titleText = t.Font.GetCapitalizedText(defaultText);
                }
            }

            if (t.Layout.HasLayout) //Only for the main chart title, axis titles don't support manual layout in Excel.
            {

                InitTextBox(maxWidth, maxHeight);
                var mr = GetRectFromManualLayout(sc, t.Layout);
                TextBox.Top = mr.Top;
                TextBox.Left = mr.Left;
            }
            else
            {
                if (axis == null)
                {
                    InitTextBox(maxWidth, maxHeight);
                    TextBox.Top = (float)6;                                       //6 point for the chart title standard offset.
                    TextBox.Left = (float)(sc.Bounds.Width - TextBox.Width) / 2;
                }
                else
                {
                    var isVertical = axis.Axis.IsVertical;
                    maxWidth = sc.Bounds.Width * (isVertical ? 0.2 : 0.8);   //Max Width.
                    maxHeight = sc.Bounds.Height * (isVertical ? 0.8 : 0.2); //Max Height.

                    InitTextBox(maxWidth, maxHeight);
                    SetAxisTitleRect(sc, axis);
                }
            }
            Rectangle.SetDrawingPropertiesFill(sc.Theme, t.Fill, sc.Chart.StyleManager.Style.Title.FillReference.Color);
            Rectangle.SetDrawingPropertiesBorder(sc.Theme, t.Border, sc.Chart.StyleManager.Style.Title.BorderReference.Color, t.Border.Fill.Style != eFillStyle.NoFill, 0.75);
        }

        private void SetAxisTitleRect(ChartRenderer sc, SvgChartAxis axis)
        {
            var margin = 8F;
            switch (axis.Axis.AxisPosition)
            {
                case eAxisPosition.Left:
                    Rectangle.Top = sc.GetPlotAreaTop();
                    Rectangle.Left = sc.Chart.HasLegend && sc.Chart.Legend.Position == eLegendPosition.Left ? sc.Legend.Rectangle.Right + LeftMargin : margin;                               
                    break;
                case eAxisPosition.Right:
                    Rectangle.Top = sc.GetPlotAreaTop();
                    Rectangle.Left = sc.Chart.HasLegend && sc.Chart.Legend.Position == eLegendPosition.Right || sc.Chart.Legend.Position == eLegendPosition.TopRight ? sc.Legend.Rectangle.Left - Rectangle.Width - margin : sc.Bounds.Right - Rectangle.Width - margin;
                    break;
                case eAxisPosition.Bottom:
                    Rectangle.Top = sc.ChartArea.Rectangle.Height - margin - Rectangle.Height;
                    Rectangle.Left = GetHorizontalLeft(sc);
                    break;
                case eAxisPosition.Top:
                    Rectangle.Top = sc.Title != null && sc.Title._title.Layout.HasLayout==false ? sc.Title.Rectangle.Bottom+margin : margin;
                    Rectangle.Left = GetHorizontalLeft(sc);
                    break;
            }
        }

        private double GetHorizontalLeft(ChartRenderer sc)
        {
            var margin = 8F;
            if (sc.HorizontalAxis!=null)
            {
                if(sc.HorizontalAxis.Axis.Deleted && sc.HorizontalAxis.Title != null && sc.HorizontalAxis.Title._title.Layout.HasLayout==false)
                {
                    return sc.HorizontalAxis.Title.Rectangle.Right;
                }
                return sc.HorizontalAxis.Rectangle.Right;
            }
            else
            {
                return margin;
            }
        }

        private static string GetDefaultChartTitleText(ChartRenderer sc, ExcelChartTitleStandard t, string defaultText)
        {
            if (string.IsNullOrEmpty(t.DisplayedText) == false)
            {
                defaultText = t.DisplayedText;
            }
            else if (sc.Chart.PlotArea.ChartTypes.Count == 1 && sc.Chart.Series.Count == 1)
            {
                if (string.IsNullOrEmpty(sc.Chart.Series[0].Header))
                {
                    var s = sc.Chart.Series[0];
                    if (s.NumberLiteralsX != null && s.NumberLiteralsX.Length > 0)
                    {
                        defaultText = s.NumberLiteralsX[0].ToString(CultureInfo.InvariantCulture);
                    }
                    else if (s.StringLiteralsX != null && s.StringLiteralsX.Length > 0)
                    {
                        defaultText = s.StringLiteralsX[0];
                    }
                    else if(s.HeaderAddress!=null)
                    {
                        defaultText = s.GetHeaderText(0);
                    }
                }
                else
                {
                    defaultText = sc.Chart.Series[0].Header;
                }
            }

            return defaultText;
        }

        internal void InitTextBox(double maxWidth, double maxHeight)
        {
            TextBox = new DrawingTextBox(_svgChart.Drawing, _svgChart.ChartArea.Rectangle.Bounds, maxWidth, maxHeight);
            if(_title.Rotation != 0)
            {
                TextBox.Rotation = _title.Rotation;
                TextBox.Rotation = _title.Rotation;
            }
            if (_title.TextBody.Paragraphs.Count > 0)
            {
                TextBox.ImportTextBody(_title.TextBody, true, ExcelHorizontalAlignment.Center);
            }
            else
            {
                var p = _title.DefaultTextBody.Paragraphs.FirstOrDefault();                
                TextBox.ImportParagraph(p, 0, _titleText);                
            }

            TextBox.LeftMargin = LeftMargin;
            TextBox.RightMargin = RightMargin;
            TextBox.TopMargin = TopMargin;
            TextBox.BottomMargin = BottomMargin;
            TextBox.TextBody.VerticalAlignment = TextAnchoringType.Top;
            Rectangle = TextBox.Rectangle;

            TextBox.LeftMargin = LeftMargin;
            TextBox.RightMargin = RightMargin;
            TextBox.TopMargin = TopMargin;
            TextBox.BottomMargin = BottomMargin;
            TextBox.TextBody.VerticalAlignment = TextAnchoringType.Top;
            Rectangle = TextBox.Rectangle;
        }

        public DrawingTextBox TextBox
        {
            get; private set;
        }
        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var p = _title.DefaultTextBody.Paragraphs.FirstOrDefault();
            TextBox.TextBody.FontColorString = "#" + p.DefaultRunProperties.Fill.Color.ToColorString();
            TextBox.Rectangle.SetDrawingPropertiesFill(_svgChart.Theme,_title.Fill, _svgChart.Chart.StyleManager.Style.Title.FillReference.Color);
            TextBox.Rectangle.SetDrawingPropertiesBorder(_svgChart.Theme, _title.Border, _svgChart.Chart.StyleManager.Style.Title.BorderReference.Color, _title.Border.Fill.Style != eFillStyle.NoFill, 0.75);
            TextBox.AppendRenderItems(renderItems);
        }

    }
}