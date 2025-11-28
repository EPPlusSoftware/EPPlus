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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartTitle : SvgChartObject
    {
        internal SvgChartTitle(SvgChart sc, ExcelChartTitleStandard t, string defaultText) : base(sc.Chart)
        {
            if (sc.Chart.HasTitle == false || sc.Chart.Series.Count == 0)
            {
                return;
            }
            //These are hard coded margins for the title box.
            LeftMargin = RightMargin = 4;
            TopMargin = BottomMargin = 2;

            var maxWidth = sc.Size.Width * 0.8;
            var maxHeight = sc.Size.Height / 2D;
            var rect = t.TextBody.Paragraphs.GetSizeInPixels(maxWidth, maxHeight, defaultText, t.Font);
           
            if (t.Layout.HasLayout)
            {
                Rectangle = GetRectFromManualLayout(sc, t.Layout);
                if(double.IsNaN(Rectangle.Width))
                {
                    Rectangle.Width = (float)(rect.Width+LeftMargin+RightMargin);
                }
                if (double.IsNaN(Rectangle.Height))
                {
                    Rectangle.Height = (float)(rect.Height+TopMargin+BottomMargin);
                }
            }
            else 
            {
                Rectangle = new SvgRenderRectItem(sc.Chart);
                Rectangle.Y = (float)8;                         //8 pixels for the chart title standard offset
                Rectangle.X = (float)(sc.Size.Width - rect.Width + LeftMargin+ RightMargin) / 2;
                Rectangle.Height = (float)(rect.Height + TopMargin + BottomMargin);
                Rectangle.Width = (float)(rect.Width + LeftMargin + RightMargin);
            }
            Rectangle.SetDrawingPropertiesFill(t.Fill, sc.Chart.StyleManager.Style.Title.FillReference.Color);
            Rectangle.SetDrawingPropertiesBorder(t.Border, sc.Chart.StyleManager.Style.Title.BorderReference.Color, t.Border.Fill.Style != eFillStyle.NoFill, 0.75);

            TextBox = new TextBox(rect.Left + LeftMargin, rect.Top + TopMargin, rect.Width - LeftMargin - RightMargin, rect.Height - RightMargin - BottomMargin);
            foreach (var p in t.TextBody.Paragraphs)
            {
                TextBox.AddParagraph(p);
            }
        }
        public TextBox TextBox
        {
            get; private set;
        }
        public override void Render(StringBuilder sb)
        {
            Rectangle.Render(sb);
            
        }
    }
}