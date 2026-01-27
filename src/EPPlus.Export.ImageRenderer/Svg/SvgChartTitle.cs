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
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Text;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartTitle : SvgChartObject
    {
        ExcelChartTitleStandard _title;
        string _titleText;
        SvgChart _svgChart;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sc"></param>
        /// <param name="t"></param>
        /// <param name="defaultText"></param>
        /// <param name="axisPosition">If null, this is the main chart title.</param>
        internal SvgChartTitle(SvgChart sc, ExcelChartTitleStandard t, string defaultText, SvgChartAxis axis=null) : base(sc)
        {
            _svgChart = sc;
            //These are hard coded margins for the title box.
            LeftMargin = RightMargin = 4;
            TopMargin = BottomMargin = 2;

            var maxWidth = sc.Bounds.Width * 0.8;
            var maxHeight = sc.Bounds.Height / 2D;
            _title = t;
            if (axis==null)
            {
                _titleText = GetDefaultChartTitleText(sc, t, defaultText);
            }
            else
            {
                if (string.IsNullOrEmpty(t.Text) == false)
                {
                    _titleText = t.Text;
                }
                else
                {
                    _titleText = defaultText;
                }
            }

            var rect = t.TextBody.Paragraphs.GetSizeInPixels(maxWidth, maxHeight, _titleText, t.Font, LeftMargin, TopMargin, t.Rotation);
            
            if (t.Layout.HasLayout)
            {
                Rectangle = GetRectFromManualLayout(sc, t.Layout);
                if (double.IsNaN(Rectangle.Width))
                {
                    Rectangle.Width = (float)(rect.Width);
                }
                if (double.IsNaN(Rectangle.Height))
                {
                    Rectangle.Height = (float)(rect.Height);
                }
                InitTextBox();
            }
            else
            {
                Rectangle = new SvgRenderRectItem(sc, sc.Bounds);
                if (axis==null)
                {   
                    Rectangle.Top = (float)8;                         //8 pixels for the chart title standard offset
                    Rectangle.Left = (float)(sc.Bounds.Width - rect.Width) / 2;
                    Rectangle.Height = (float)rect.Height;
                    Rectangle.Width = (float)rect.Width;
                    InitTextBox();
                }
                else
                {
                    SetAxisTitleRect(sc, axis, rect);
                }
            }

            Rectangle.SetDrawingPropertiesFill(t.Fill, sc.Chart.StyleManager.Style.Title.FillReference.Color);
            Rectangle.SetDrawingPropertiesBorder(t.Border, sc.Chart.StyleManager.Style.Title.BorderReference.Color, t.Border.Fill.Style != eFillStyle.NoFill, 0.75);
        }

        private void SetAxisTitleRect(SvgChart sc, SvgChartAxis axis, RectBase rect)
        {
            var margin = 8F;
            switch (axis.Axis.AxisPosition)
            {
                case eAxisPosition.Left:
                    Rectangle.Top = sc.GetPlotAreaTop();
                    Rectangle.Left = sc.Chart.HasLegend && sc.Chart.Legend.Position == eLegendPosition.Left ? sc.Legend.Rectangle.Right : margin;                               
                    break;
                case eAxisPosition.Bottom:
                    Rectangle.Top = sc.ChartArea.Height - margin - rect.Height;
                    Rectangle.Left = GetHorizontalLeft(sc);
                    break;
            }
            Rectangle.Width = rect.Width;
            Rectangle.Height = rect.Height;
        }

        private double GetHorizontalLeft(SvgChart sc)
        {
            var margin = 8F;
            if (sc.HorizontalAxis!=null)
            {
                return sc.HorizontalAxis.Rectangle.Right;
            }
            else if(sc.HorizontalAxisTitle != null)
            {
                return sc.HorizontalAxisTitle.Rectangle.Right;
            }
            else
            {
                return margin;
            }
        }

        private static string GetDefaultChartTitleText(SvgChart sc, ExcelChartTitleStandard t, string defaultText)
        {
            if (string.IsNullOrEmpty(t.Text) == false)
            {
                defaultText = t.Text;
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
                    else
                    {
                        defaultText = s.GetHeaderText();
                    }
                }
                else
                {
                    defaultText = sc.Chart.Series[0].Header;
                }
            }

            return defaultText;
        }

        internal void InitTextBox()
        {
            TextBox = new SvgTextBoxItem(_svgChart, _svgChart.ChartArea.Bounds, Rectangle.Left, Rectangle.Top, Rectangle.Width, Rectangle.Height);
            TextBox.LeftMargin = LeftMargin;
            TextBox.RightMargin = RightMargin;
            TextBox.TopMargin = TopMargin;
            TextBox.BottomMargin = BottomMargin;
            TextBox.TextBody.VerticalAlignment = eTextAnchoringType.Top;
            if(_title.Rotation != 0)
            {
                TextBox.Bounds.Rotation = _title.Rotation;
            }
            if (_title.TextBody.Paragraphs.Count > 0)
            {
                foreach (var p in _title.TextBody.Paragraphs)
                {
                    TextBox.TextBody.ImportParagraph(p, 0);
                }
            }
            else
            {
                //TextBox.AddText(string.IsNullOrEmpty(_title.Text) ? _titleText : _title.Text, _title.Font);
                var text = string.IsNullOrEmpty(_title.Text) ? _titleText : _title.Text;
                var p = _title.DefaultTextBody.Paragraphs.FirstOrDefault();                
                TextBox.TextBody.ImportParagraph(p, 0, text);
            }
        }

        public SvgTextBoxItem TextBox
        {
            get; private set;
        }
        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {

            TextBox.AppendRenderItems(renderItems);
            TextBox.Rectangle.SetDrawingPropertiesFill(_title.Fill, _svgChart.Chart.StyleManager.Style.Title.FillReference.Color);
            TextBox.Rectangle.SetDrawingPropertiesBorder(_title.Border, _svgChart.Chart.StyleManager.Style.Title.BorderReference.Color, _title.Border.Fill.Style != eFillStyle.NoFill, 0.75);
        }

    }
}