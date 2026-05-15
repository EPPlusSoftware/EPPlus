using EPPlusImageRenderer.FontData;
using EPPlusImageRenderer.RenderItems.Shared;
using EPPlusImageRenderer.Svg;
using EPPlusImageRenderer.Text;
using EPPlusImageRenderer.Utils;
using Microsoft.Extensions.Primitives;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.Text;


namespace EPPlusImageRenderer.RenderItems.Svg
{
    internal class SvgParagraph : SvgRenderItem
    {
        public override void Render(StringBuilder sb)
        {
            sb.Append("<text");
            base.Render(sb);

            sb.Append($" {vertAlignAttribute} " + GetHorizontalAlignmentAttribute(XPos) +
                $"font-family=\"{fontName},{fontName}_MSFontService,sans-serif\" " +
                $"font-size=\"{(defaultMeasureFont.Size / 16).ToString(CultureInfo.InvariantCulture)}em\">");

            foreach (var textRun in TextRuns)
            {
               textRun.Render(sb);
            }

            sb.Append("</text>");
        }

        public override SvgItemType Type => SvgItemType.Text;

        internal override RenderItem Clone(SvgShape svgDocument)
        {
            throw new NotImplementedException();
        }

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = ParagraphArea.Left + LeftMargin;
            it = ParagraphArea.Right - RightMargin;
            ir = ParagraphArea.Top;
            ib = ParagraphArea.Bottom;
        }

        string vertAlignAttribute = "";
        protected List<SvgTextRun> TextRuns = new List<SvgTextRun>();

        //ExcelDrawingParagraph Paragraph;
        double RightMargin;
        double LeftMargin;

        eTextAlignment HorizontalAlignment;

        RectBase ParagraphArea;

        ExcelDrawingParagraph Paragraph;

        /// <summary>
        /// X position
        /// </summary>
        protected double XPos;

        string fontName;

        /// <summary>
        /// Line-Spacing in pixels
        /// </summary>
        protected double LineSpacing;
        MeasurementFont defaultMeasureFont;

        public SvgParagraph(ExcelDrawingParagraph p, RectBase paragraphArea, FontMeasurerExact fmExact, string VertAlign)
        {
            Paragraph = p;
            defaultMeasureFont = p.DefaultRunProperties.GetMeasureFont();
            vertAlignAttribute = VertAlign;
            //Seperated out in case of some final render item
            //needing to adjust without changing the original values
            RightMargin = p.RightMargin;
            LeftMargin = p.LeftMargin;



            ParagraphArea = paragraphArea;
            HorizontalAlignment = p.HorizontalAlignment;

            LineSpacing = GetParagraphLineSpacingInPixels(fmExact);
            XPos = GetAlignmentHorizontal(HorizontalAlignment);
        }

        private string GetHorizontalAlignmentAttribute(double indentX)
        {
            string ret = "";
            var xStr = indentX.ToString(CultureInfo.InvariantCulture);

            switch (HorizontalAlignment)
            {
                default:
                case eTextAlignment.Left:
                    ret = $"text-anchor=\"start\" x=\"{xStr}\" ";
                    break;
                case eTextAlignment.Center:
                    ret = $"text-anchor=\"middle\" x=\"{xStr}\" ";
                    break;
                case eTextAlignment.Right:

                    ret = $"text-anchor=\"end\" x=\"{xStr}\" ";
                    break;
            }

            return ret;
        }

        private double GetParagraphLineSpacingInPixels(FontMeasurerExact fmExact)
        {
            if (Paragraph.LineSpacing.LineSpacingType == eDrawingTextLineSpacing.Exactly)
            {
                return Paragraph.LineSpacing.Value.PointToPixel();
            }
            else
            {
                var multiplier = (Paragraph.LineSpacing.Value / 100);
                return multiplier * fmExact.GetSingleLineSpacingPX();
            }
        }

        internal double GetAlignmentHorizontal(eTextAlignment txAlignment)
        {
            var area = ParagraphArea;
            double x = 0;
            switch (txAlignment)
            {
                case eTextAlignment.Left:
                default:
                    x = area.Left;
                    break;
                case eTextAlignment.Center:
                    x = (area.Right / 2) + LeftMargin - RightMargin;
                    break;
                case eTextAlignment.Right:
                    x = area.Right - RightMargin;
                    break;
            }

            return TextUtils.RoundToWhole(x);
        }

        internal void AddTextRun(ExcelParagraphTextRunBase txtRun)
        {
            GetBounds(out double l, out double t, out double r, out double b);
            var textMaxWidth = r - l;
            var textRun = new SvgTextRun(txtRun, LineSpacing, textMaxWidth);
            TextRuns.Add(textRun);
        }
    }
}
