using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.Svg.NodeAttributes;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgParagraphItem : ParagraphContainer
    {
        public SvgParagraphItem()
        {
        }

        public SvgParagraphItem(BoundingBox parent) : base(parent)
        {
        }

        public SvgParagraphItem(ExcelDrawingParagraph p, BoundingBox parent) : base(p, parent)
        {
            //---Add Actual textruns---
            for (int i = 0; i < p.TextRuns.Count; i++)
            {
                AddTextRun(p.TextRuns[i], _textRunContent[i]);
            }
        }

        internal protected override void AddTextRun(ExcelParagraphTextRunBase origTxtRun, string runDisplayString)
        {
            AddTextRun<SvgTextRunItem>(origTxtRun, runDisplayString);
        }

        private string GetHorizontalAlignmentAttribute(double indentX)
        {
            string ret = "";
            var xStr = indentX.ToString(CultureInfo.InvariantCulture);

            switch (_hAlign)
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


        //public void BaseRender(StringBuilder sb)
        //{
        //    if (string.IsNullOrEmpty(FillColor) == false)
        //    {
        //        sb.Append($"fill=\"{FillColor}\" ");
        //        if (FillOpacity != null && FillOpacity != 1)
        //        {
        //            sb.Append($"opacity=\"{FillOpacity.Value.ToString(CultureInfo.InvariantCulture)}\" ");
        //        }
        //    }
        //    if (string.IsNullOrEmpty(FilterName) == false)
        //    {
        //        sb.Append($"filter=\"{FilterName}\" ");
        //    }

        //    if (BorderWidth.HasValue && string.IsNullOrEmpty(BorderColor) == false)
        //    {
        //        sb.Append($"stroke=\"{BorderColor}\" ");
        //    }
        //    if (BorderWidth.HasValue)
        //    {
        //        var v = BorderWidth.Value * ExcelDrawing.EMU_PER_POINT / ExcelDrawing.EMU_PER_PIXEL;
        //        sb.Append($"stroke-width=\"{v.ToString(CultureInfo.InvariantCulture)}\" ");

        //        if (BorderDashArray != null)
        //        {
        //            var BorderDashArrayStr = BorderDashArray.Select(x =>
        //            x.ToString(CultureInfo.InvariantCulture)).ToArray();

        //            sb.Append($"stroke-dasharray=\"" + $"{string.Join(",", BorderDashArrayStr)}\" ");
        //        }
        //    }

        //    sb.Append($"stroke-miterlimit =\"8\" ");
        //}

        public override void Render(StringBuilder sb)
        {
            var fontSize = _paragraphFont.Size.PointToPixel().ToString(CultureInfo.InvariantCulture);

            sb.AppendLine($"<g transform=\"translate({Bounds.X},{Bounds.Y})\" >");

            sb.AppendLine("<title>");
            sb.AppendLine("paragraph");
            sb.AppendLine("</title>");


            sb.Append("<text ");
            base.Render(sb);

            sb.Append(/*$"{GetHorizontalAlignmentAttribute(Bounds.X)} y=\"{Bounds.Y}\" " +*/
                $"font-family=\"{_paragraphFont.FontFamily},{_paragraphFont.FontFamily}_MSFontService,sans-serif\" " +
                $"font-size=\"{fontSize}px\" >");

            if (_textRunItems.Count > 0)
            {
                foreach (var textRun in _textRunItems)
                {
                    textRun.Render(sb);
                }
            }

            sb.AppendLine("</text>");
            sb.AppendLine("</g>");

            //var paragraphGroup = new SvgElement("g");
            //paragraphGroup.AddAttribute("transform", $"translate({Bounds.X},{Bounds.Y})");

            //var paragraphTitle = new SvgElement("title");
            //paragraphTitle.Content = "paragraph";
            ////paragraphTitle.Content = "Paragraph " + paragraphCount.ToString();
            //paragraphGroup.AddChildElement(paragraphTitle);

            //var paragraphElement = new SvgElement("text");
            //paragraphElement.AddAttribute("y", _paragraphFont.Size.PointToPixel());
            //paragraphElement.AddAttribute("font-size", $"{_paragraphFont.Size.PointToPixel()}px");
            ////paragraphElement.AddAttribute("clip-path", $"url(#{nameId})");

            ////paragraphGroup.AddChildElement(paragraphElement);

            //sb.AppendLine(RenderSvgElement(paragraphElement));

            //foreach (var run in _textRunItems)
            //{
            //    run.Render(sb);
            //}

            //foreach(var run in _textRunItems)
            //{
            //    var bbVisual = new SvgElement("rect");
            //    bbVisual.AddAttribute("x", run.Bounds.X);
            //    bbVisual.AddAttribute("y", run.Bounds.Y);
            //    bbVisual.AddAttribute("width", run.Bounds.Width);
            //    bbVisual.AddAttribute("height", run.Bounds.Height);
            //    bbVisual.AddAttribute("fill", "blue");
            //    bbVisual.AddAttribute("opacity", "0.5");
            //    sb.AppendLine(RenderSvgElement(bbVisual));
            //    sb.AppendLine("</rect>");
            //}

            //sb.AppendLine("</g>");
        }

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            throw new System.NotImplementedException();
        }

        //public override void Render(SvgElement textBodyGroup, int paragraphCount, string nameId)
        //{
        //    var paragraphGroup = new SvgElement("g");
        //    paragraphGroup.AddAttribute("transform", $"translate({Bounds.X},{Bounds.Y})");

        //    var paragraphTitle = new SvgElement("title");
        //    paragraphTitle.Content = "Paragraph " + paragraphCount.ToString();
        //    paragraphGroup.AddChildElement(paragraphTitle);

        //    var paragraphElement = new SvgElement("text");
        //    paragraphElement.AddAttribute("y", _paragraphFont.Size.PointToPixel());
        //    paragraphElement.AddAttribute("font-size", $"{_paragraphFont.Size.PointToPixel()}px");
        //    paragraphElement.AddAttribute("clip-path", $"url(#{nameId})");

        //    paragraphGroup.AddChildElement(paragraphElement);


        //    textBodyGroup.AddChildElement(paragraphGroup);
        //}
    }
}
