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
using EPPlusImageRenderer.Svg;
using EPPlusImageRenderer.Text;
using OfficeOpenXml.Drawing;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using EPPlus.Graphics;

namespace EPPlusImageRenderer.RenderItems
{
    internal class SvgRenderRectItem : SvgRenderItem
    {
        public SvgRenderRectItem() : base()
        {

        }

        public SvgRenderRectItem(ExcelDrawing drawing) : base(drawing)
        {

        }
        //List<SvgParagraph> Paragraphs = new List<SvgParagraph>();
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double Bottom => Top + Height;
        public double Right => Left + Width;
        public override SvgItemType Type => SvgItemType.Rect;

        public override void Render(StringBuilder sb)
        {
            //if (Paragraphs.Count == 0)
            //{
                RenderRect(sb);
            //}
            //else
            //{
                //var groupItem = new SvgGroupItem("");
                //groupItem.Render(sb);

                //RenderRect(sb);

                ////TODO: add textbody property stuff here
                //foreach (var paragraph in Paragraphs)
                //{
                //    paragraph.Render(sb);
                //}
                //groupItem.RenderEndGroup(sb);
            //}
        }

        private void RenderRect(StringBuilder sb)
        {
            sb.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" ",
                Left.ToString(CultureInfo.InvariantCulture),
                Top.ToString(CultureInfo.InvariantCulture),
                Width.ToString(CultureInfo.InvariantCulture),
                Height.ToString(CultureInfo.InvariantCulture));
            base.Render(sb);
            sb.AppendFormat("/>");
        }

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            var clone = new SvgRenderRectItem(_drawing);
            clone._theme = _theme;
            CloneBase(clone);

            clone.Left = Left * svgDocument.Size.Width;
            clone.Top = Top * svgDocument.Size.Height;
            clone.Width = svgDocument.Size.Width * Width;
            clone.Height = svgDocument.Size.Height * Height;

            return clone;
        }
        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = Left;
            it = Top;
            ir = Width;
            ib = Height;
        }
        //internal string fontColor;
        //internal eTextAnchoringType VerticalAlignment;

        //internal SvgParagraph AddParagraph(ExcelDrawingParagraph item)
        //{            
        //    var measureFont = item._paragraphs[0].DefaultRunProperties.GetMeasureFont();

        //    //Document Top position for the paragraph text based on vertical alignment
        //    var posY = GetAlignmentVertical();
        //    var vertAlignAttribute = GetVerticalAlignAttribute(posY);

        //    //var area = Bounds.GetInnerRect();

        //    //Limit bounding area with the space taken by previous paragraphs
        //    //Note that this is ONLY identical to PosY if the vertical alignment is top
        //    //area.Top = paragraphStartPosY;

        //    //The first run in the first paragraph must apply different line-spacing
        //    bool isFirst = Paragraphs.Count == 0;
        //    var svgParagraph = new SvgParagraph(item, area, vertAlignAttribute, posY, isFirst);

        //    svgParagraph.FillColor = string.IsNullOrEmpty(fontColor) ? item.DefaultRunProperties.Fill.Color.Name : fontColor;

        //    paragraphStartPosY = svgParagraph.GetBottomYPosition();

        //    Paragraphs.Add(svgParagraph);
        //    return svgParagraph;
        //}
        /// <summary>
        /// Get the start of text space vertically
        /// </summary>
        /// <param name="fontSizeInPixels"></param>
        /// <returns></returns>
        //internal double GetAlignmentVertical()
        //{
        //    double alignmentY = 0;

        //    var y = paragraphStartPosY;

        //    var height = y - Bounds.Bottom;

        //    switch (VerticalAlignment)
        //    {
        //        case eTextAnchoringType.Top:
        //            alignmentY = y;
        //            break;
        //        case eTextAnchoringType.Center:
        //            var adjustedHeight = y + (height / 2);

        //            alignmentY = adjustedHeight + Bounds.MarginTop;
        //            break;
        //        case eTextAnchoringType.Bottom:
        //            alignmentY = height - Bounds.MarginBottom;
        //            break;
        //    }

        //    return alignmentY;
        //}

        //private string GetVerticalAlignAttribute(double textOrigY)
        //{
        //    string ret = "dominant-baseline=";

        //    switch (VerticalAlignment)
        //    {
        //        case eTextAnchoringType.Top:
        //            ret += $"\"text-top\" ";
        //            break;
        //        case eTextAnchoringType.Center:
        //            ret += $"\"middle\" ";
        //            break;
        //        case eTextAnchoringType.Bottom:
        //            ret += $"\"text-bottom\" ";
        //            break;
        //        default:
        //            return "";
        //    }

        //    var yStrValue = textOrigY.ToString(CultureInfo.InvariantCulture);
        //    ret += $" y=\"{yStrValue}\"";

        //    return ret;
        //}

    }
}