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
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Text;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Globalization;
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartTitle : SvgChartObject
    {
        internal SvgChartTitle(SvgChart sc, ExcelChartTitleStandard t, string defaultText, bool isAxis) : base(sc.Chart)
        {
            if (isAxis==false && sc.Chart.HasTitle == false || sc.Chart.Series.Count == 0)
            {
                return;
            }
            //These are hard coded margins for the title box.
            LeftMargin = RightMargin = 4;
            TopMargin = BottomMargin = 2;

            var maxWidth = sc.Size.Width * 0.8;
            var maxHeight = sc.Size.Height / 2D;
            if(string.IsNullOrEmpty(t.Text)==false)
            {
                defaultText = t.Text;
            }
            else if(sc.Chart.PlotArea.ChartTypes.Count == 1 && sc.Chart.Series.Count == 1)
            {
                if (string.IsNullOrEmpty(sc.Chart.Series[0].Header))
                {
                    var s=sc.Chart.Series[0];
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

            var rect = t.TextBody.Paragraphs.GetSizeInPixels(maxWidth, maxHeight, defaultText, t.Font);

            if (t.Layout.HasLayout)
            {
                Rectangle = GetRectFromManualLayout(sc, t.Layout);
                if (double.IsNaN(Rectangle.Width))
                {
                    Rectangle.Width = (float)(rect.Width + LeftMargin + RightMargin);
                }
                if (double.IsNaN(Rectangle.Height))
                {
                    Rectangle.Height = (float)(rect.Height + TopMargin + BottomMargin);
                }
            }
            else
            {
                Rectangle = new SvgRenderRectItem(sc.Chart);
                Rectangle.Y = (float)8;                         //8 pixels for the chart title standard offset
                Rectangle.X = (float)(sc.Size.Width - rect.Width + LeftMargin + RightMargin) / 2;
                Rectangle.Height = (float)(rect.Height + TopMargin + BottomMargin);
                Rectangle.Width = (float)(rect.Width + LeftMargin + RightMargin);
            }
            Rectangle.SetDrawingPropertiesFill(t.Fill, sc.Chart.StyleManager.Style.Title.FillReference.Color);
            Rectangle.SetDrawingPropertiesBorder(t.Border, sc.Chart.StyleManager.Style.Title.BorderReference.Color, t.Border.Fill.Style != eFillStyle.NoFill, 0.75);

            InitTextBox(t, defaultText);
        }

        private void InitTextBox(ExcelChartTitleStandard t, string defaultText)
        {
            TextBox = new TextBox(Chart, Rectangle.X, Rectangle.Y , Rectangle.Width, Rectangle.Height);
            TextBox.Bounds.MarginLeft = LeftMargin;
            TextBox.Bounds.MarginRight = RightMargin;
            TextBox.Bounds.MarginTop = TopMargin;
            TextBox.Bounds.MarginBottom = BottomMargin;
            TextBox.VerticalAlignment = eTextAnchoringType.Top;
            if (t.TextBody.Paragraphs.Count > 0)
            {
                foreach (var p in t.TextBody.Paragraphs)
                {
                    TextBox.AddParagraph(p);
                }
            }
            else
            {
                TextBox.AddText(string.IsNullOrEmpty(t.Text) ? defaultText : t.Text, t.Font);
            }
        }

        public TextBox TextBox
        {
            get; private set;
        }
        public override void Render(StringBuilder sb)
        {
            Rectangle.Render(sb);
            TextBox.Render(sb);
        }
    }
}