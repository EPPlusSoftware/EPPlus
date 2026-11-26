using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.RenderItems.SvgItem;
using EPPlus.DrawingRenderer.Utils;
using EPPlus.Export.ImageRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;
using OfficeOpenXml.Utils;
using System.Globalization;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    internal class SvgTextRunRenderer : SvgBaseRenderer<TextRunRenderItem>
    {
        string UnderlineColorString = string.Empty;

        public SvgTextRunRenderer(StringBuilder outputStream) : base(outputStream)
        {

        }

        internal SvgFontPlan FontPlan { get; set; } = SvgFontPlan.WithoutContext;

        /// <summary>
        /// Appends the font-style, font-weight and text-decoration declarations of the run. Weight and
        /// style follow the subfamily the text was measured with, which is also the subfamily of the
        /// embedded face, so output cannot diverge from measurement; the run's own flags are used only
        /// when it has no measurement font. When the underline has a color of its own, text-decoration
        /// goes on a wrapping tspan instead (see <see cref="UnderlineColorString"/>), whose fill is the
        /// color the line is drawn with.
        /// </summary>
        void AppendFontStyleDeclarations(TextRunRenderItem textRun, StringBuilder style)
        {
            bool bold, italic;
            if (textRun._measurementFont != null)
            {
                var subFamily = textRun._measurementFont.SubFamily;
                bold = subFamily == FontSubFamily.Bold || subFamily == FontSubFamily.BoldItalic;
                italic = subFamily == FontSubFamily.Italic || subFamily == FontSubFamily.BoldItalic;
            }
            else
            {
                bold = textRun._isBold;
                italic = textRun._isItalic;
            }

            if (italic)
            {
                style.Append("font-style:italic;");
            }
            if (bold)
            {
                style.Append("font-weight:bold;");
            }

            string decoration = string.Empty;
            if (textRun._underLineType != eDrawingUnderLineType.None)
            {
                // Double, dotted, dashed and wavy underlines need text-decoration-style, which svg
                // renderers do not support consistently, so every type is drawn as a single underline.
                decoration = "underline";
                if (textRun._underlineColor != System.Drawing.Color.Empty)
                {
                    UnderlineColorString = $"<tspan style=\"fill:#{textRun._underlineColor.To6CharHexString()};{{0}}\">{{1}}</tspan>";
                }
            }
            if (textRun._strikeType == eDrawingStrikeType.Single)
            {
                // Line values are separated by spaces; a comma makes the whole declaration invalid.
                decoration += decoration.Length > 0 ? " line-through" : "line-through";
            }

            if (decoration.Length == 0)
            {
                return;
            }
            if (string.IsNullOrEmpty(UnderlineColorString))
            {
                style.Append("text-decoration:").Append(decoration).Append(';');
            }
            else
            {
                UnderlineColorString = string.Format(UnderlineColorString, $"text-decoration:{decoration};", "{0}");
            }
        }

        /// <summary>
        /// The textLength and lengthAdjust attributes that make the browser draw the text with the
        /// width it was measured with, whatever font it ends up using. Empty when the text has no
        /// measured width.
        /// </summary>
        string GetTextLengthAttributes(IFontFormatBase font, string text)
        {
            var width = FontPlan.MeasureWidthInPoints(font, text);
            if (width.HasValue == false || width.Value <= 0)
            {
                return string.Empty;
            }
            return $"textLength=\"{width.Value.PointToPixelString()}\" lengthAdjust=\"spacing\" ";
        }

        public override void Render(TextRunRenderItem textRun)
        {
            double fontSize = textRun.FontSizeInPixels;

            // Text measured with a glyph fallback written by name is split into child tspans, each with
            // its own font-family and textLength. Otherwise the run carries textLength itself.
            var segments = FontPlan.GetFallbackSegments(textRun._measurementFont, textRun.CurrentText);

            var sb = new StringBuilder();
            sb.Append("<tspan ");
            if (textRun._baseline != 0)
            {
                // For sub- and superscript, move the text up or down by baseline% of the font size.
                // A negative value moves it up, a positive value down.
                sb.Append($"dy=\"{(fontSize.PointToPixel() * -textRun._baseline / 100D).ToString(CultureInfo.InvariantCulture)}px\" ");
            }
            sb.Append($"x=\"{textRun.Left.PointToPixelString()}\" ");
            sb.Append($"y=\"{textRun.YPosition.PointToPixelString()}px\" ");

            if (textRun._measurementFont != null && segments == null)
            {
                sb.Append(GetTextLengthAttributes(textRun._measurementFont, textRun.CurrentText));
            }

            // Font and fill go in the style attribute rather than in presentation attributes: an inline
            // style wins over the css rules of a host page (e.g. "* { font-family: ... }"), which beat
            // presentation attributes and would draw the text with another font than it was measured with.
            var style = new StringBuilder();
            if (textRun._measurementFont != null)
            {
                style.Append("font-family:").Append(FontPlan.GetFontFamilyAttributeValue(textRun._measurementFont)).Append(';');
                style.Append("font-size:").Append(fontSize.ToString(CultureInfo.InvariantCulture)).Append("px;");
            }
            AppendFontStyleDeclarations(textRun, style);
            if (string.IsNullOrEmpty(textRun.Style.FillColor) == false)
            {
                style.Append("fill:").Append(textRun.Style.FillColor).Append(';');
            }
            if (style.Length > 0)
            {
                sb.Append("style=\"").Append(style).Append("\" ");
            }

            // The fill is in the style attribute, so the base must not write it as an attribute as well.
            var fillColor = textRun.Style.FillColor;
            textRun.Style.FillColor = null;
            RenderBaseToSpecified(textRun, sb);
            textRun.Style.FillColor = fillColor;

            sb.Append('>');
            if (segments == null)
            {
                sb.Append(ConvertUtil.HtmlEscapeString(textRun.CurrentText));
            }
            else
            {
                foreach (var segment in segments)
                {
                    sb.Append("<tspan ");
                    if (segment.FontFamily != null)
                    {
                        sb.Append("style=\"font-family:").Append(segment.FontFamily).Append(";\" ");
                    }
                    sb.Append(GetTextLengthAttributes(textRun._measurementFont, segment.Text));
                    sb.Append('>').Append(ConvertUtil.HtmlEscapeString(segment.Text)).Append("</tspan>");
                }
            }
            sb.Append("</tspan>");

            var textRunString = sb.ToString();

            // Wrap in another tspan to draw the underline in its own color.
            if (string.IsNullOrEmpty(UnderlineColorString) == false)
            {
                textRunString = string.Format(UnderlineColorString, textRunString);
            }

            if (textRun is SvgTextRunRenderItem tr && tr.RenderTextNode)
            {
                textRunString = $"<text xml:space=\"preserve\">{textRunString}</text>";
            }

            OutputStream.Append(textRunString);
            UnderlineColorString = string.Empty;
        }
    }
}