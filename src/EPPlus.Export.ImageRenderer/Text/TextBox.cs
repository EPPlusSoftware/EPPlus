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
using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace EPPlusImageRenderer.Text
{
    internal class TextBox : RenderItem
    {
        internal RectMargins Bounds = new RectMargins();

        internal eTextAnchoringType VerticalAlignment;

        internal double ClippingHeight;

        internal bool WrapText = true;

        internal double Width
        {
            get;
            set;
        }

        internal double Height
        {
            get;
            set;
        }

        public override SvgItemType Type => SvgItemType.Rect;

        /// <summary>
        /// Top of the paragraph bounding box
        /// </summary>
        double paragraphStartPosY = 0;
        private OfficeOpenXml.Interfaces.Drawing.Text.MeasurementFont mFontTextRun = null;

        List<SvgParagraph> Paragraphs = new List<SvgParagraph>();
        List<SvgTextRun> CellTextRuns = new List<SvgTextRun>();

        internal string fontColor;

        /// <summary>
        /// The input parameters are assumed to be in pixels
        /// </summary>
        /// <param name="l"></param>
        /// <param name="t"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        internal TextBox(double l, double t, double width, double height)
        {
            Bounds.Left = l; Bounds.Top = t; Width = width; Height = height;

            Bounds.Right = Bounds.Left + Width;
            Bounds.Bottom = Bounds.Top + Height;

            //Initalize margins to 0
            Bounds.MarginLeft = Bounds.MarginRight = Bounds.MarginTop = Bounds.MarginBottom = 0;

            paragraphStartPosY = Bounds.Top;

            ClippingHeight = Bounds.GetInnerBottom();
        }

        /// <summary>
        /// It is pressumed that txBody has units in EMUs and SvgRenderRectItem in pixels
        /// </summary>
        /// <param name="txBody"></param>
        /// <param name="textBoxRender"></param>
        internal TextBox(ExcelTextBody txBody, SvgRenderRectItem textBoxRender)
        {
            Bounds.Left = textBoxRender.X;
            Bounds.Top = textBoxRender.Y;

            Width = textBoxRender.Width;
            Height = textBoxRender.Height;

            Bounds.Right = Bounds.Left + Width;
            Bounds.Bottom = Bounds.Top + Height;

            txBody.GetInsetsOrDefaults(out double l, out double r, out double t, out double b);

            //Ensure margins are in pixels
            Bounds.MarginLeft = l.PointToPixel();
            Bounds.MarginRight = r.PointToPixel();
            Bounds.MarginTop = t.PointToPixel();
            Bounds.MarginBottom = b.PointToPixel();

            paragraphStartPosY = Bounds.Top + Bounds.MarginTop;

            ClippingHeight = Bounds.GetInnerBottom();
        }

        public TextBox()
        {
        }

        double? _alignmentY = null;

        /// <summary>
        /// Get the start of text space vertically
        /// </summary>
        /// <param name="fontSizeInPixels"></param>
        /// <returns></returns>
        private double GetAlignmentVertical()
        {
            double alignmentY = 0;

            var y = paragraphStartPosY;

            var height = y - Bounds.Bottom;

            switch (VerticalAlignment)
            {
                case eTextAnchoringType.Top:
                    alignmentY = y;
                    break;
                case eTextAnchoringType.Center:
                    var adjustedHeight = y + (height / 2);

                    alignmentY = adjustedHeight + Bounds.MarginTop;
                    break;
                case eTextAnchoringType.Bottom:
                    alignmentY = height - Bounds.MarginBottom;
                    break;
            }

            _alignmentY = alignmentY;

            return _alignmentY.Value;
        }

        private string GetVerticalAlignAttribute(double textOrigY)
        {
            string ret = "dominant-baseline=";

            switch (VerticalAlignment)
            {
                case eTextAnchoringType.Top:
                    ret += $"\"text-top\" ";
                    break;
                case eTextAnchoringType.Center:
                    ret += $"\"middle\" ";
                    break;
                case eTextAnchoringType.Bottom:
                    ret += $"\"text-bottom\" ";
                    break;
                default:
                    return "";
            }

            var yStrValue = textOrigY.ToString(CultureInfo.InvariantCulture);
            ret += $" y=\"{yStrValue}\"";

            return ret;
        }

        public RectBase GetTextArea()
        {
            return Bounds.GetInnerRect();
        }

        internal SvgParagraph AddParagraph(ExcelDrawingParagraph item)
        {
            var measureFont = item.DefaultRunProperties.GetMeasureFont();

            //Document Y position for the paragraph text based on vertical alignment
            var posY = GetAlignmentVertical();
            var vertAlignAttribute = GetVerticalAlignAttribute(posY);

            var area = GetTextArea();

            //Limit bounding area with the space taken by previous paragraphs
            //Note that this is ONLY identical to PosY if the vertical alignment is top
            area.Top = paragraphStartPosY;

            //The first run in the first paragraph must apply different line-spacing
            bool isFirst = Paragraphs.Count == 0;
            var svgParagraph = new SvgParagraph(item, area, vertAlignAttribute, posY, isFirst);

            svgParagraph.FillColor = string.IsNullOrEmpty(fontColor) ? item.DefaultRunProperties.Fill.Color.Name : fontColor;

            paragraphStartPosY = svgParagraph.GetBottomYPosition();

            Paragraphs.Add(svgParagraph);
            return svgParagraph;
        }

        internal void RemoveParagraph(SvgParagraph item)
        {
            if(Paragraphs.Contains(item))
            {
                Paragraphs.Remove(item);
            }
        }
        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = Bounds.GetInnerLeft();
            ir = Bounds.GetInnerRight();
            it = Bounds.GetInnerTop();
            ib = Bounds.GetInnerBottom();
        }

        public override void Render(StringBuilder sb)
        {
            RenderParagraphs(sb);
        }

        internal void RenderParagraphs(StringBuilder sb)
        {
            var groupItem = new SvgGroupItem("");
            groupItem.Render(sb);
            //TODO: add textbody property stuff here
            foreach (var paragraph in Paragraphs)
            {
                paragraph.Render(sb);
            }
            groupItem.RenderEndGroup(sb);
        }
        internal void RenderTextRuns(StringBuilder sb)
        {
            var groupItem = new SvgGroupItem("");
            groupItem.Render(sb);

            var innerTop = Bounds.GetInnerRect().Top;
            var fontFamilyAttr = $"font-family=\"{mFontTextRun.FontFamily},{mFontTextRun.FontFamily}_MSFontService,sans-serif\" ";
            sb.Append($"<text fill=\"black\" y=\"{innerTop.ToString(CultureInfo.InvariantCulture)}\" {fontFamilyAttr}>");
            //TODO: add textbody property stuff here
            foreach (var textRun in CellTextRuns)
            {
                textRun.Render(sb);
            }
            sb.Append("</text>");
            groupItem.RenderEndGroup(sb);
        }

        internal void AddCellTextRun(ExcelRangeBase cell)
        {
            var fontStyle = cell.Style.Font;

            //Line spacing in points and left margin in pixels
            double lineSpacing = (0.205d * fontStyle.Size + 1);
            Bounds.MarginLeft = lineSpacing;
            Bounds.MarginRight = lineSpacing;

            mFontTextRun = fontStyle.GetMeasureFont();

            var lineSpacingPixels = (lineSpacing / 72d) * 92d;

            var posY = GetAlignmentVertical();
            var vertAlignAttribute = GetVerticalAlignAttribute(posY);

            var area = GetTextArea();
            var horizontalAlign = cell.Style.HorizontalAlignment;


            foreach (var rt in cell.RichText)
            {
                var textrun = new SvgTextRun(rt, lineSpacingPixels,
                    Bounds.GetInnerRight(), Bounds.GetInnerBottom(),
                    Bounds.GetInnerLeft(), Bounds.GetInnerTop(), mFontTextRun, horizontalAlign);

                CellTextRuns.Add(textrun);
            }

            //var lines = CellTextRuns.Last().GetLineCount();

            //if(lines * )

            //if (botPos > Bounds.Height)
            //{
            //    //Negative values increases bounds
            //    Bounds.Bottom = Bounds.Top + botPos;
            //}
        }
    }
}
