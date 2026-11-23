/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  23/11/2026         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Subsetting;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;
using System.Globalization;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    /// <summary>
    /// Decides the fonts of one svg: the CSS font-family stacks, and, when fonts are embedded, the
    /// subsets declared with @font-face.
    ///
    /// The text of the svg is collected first (<see cref="BeginCollect"/>, <see cref="AddText"/>,
    /// <see cref="Build"/>) and a <see cref="WebFontSubsetPlanner"/> decides the stacks, and the subsets
    /// when fonts are embedded. A stack then starts with the subsets, e.g.
    /// <c>'ABCDEF+Aptos Narrow', 'Aptos Narrow', 'Calibri', ..., sans-serif</c>.
    ///
    /// When fonts are not embedded, a stack is the font measurement used, the glyph fallbacks it used
    /// (each followed by its script's fallback chain), the measured family and its fallback chain
    /// (user-configured, then built-in), and a generic family, e.g.
    /// <c>'Calibri', 'Carlito', 'Liberation Sans', 'Arial', 'Helvetica', sans-serif</c>.
    /// The measured family is the family after web font substitution, so it is the family the
    /// text was laid out for. Text measured with a glyph fallback that is written by name is split off
    /// with a stack of its own (<see cref="GetFallbackSegments"/>), whether fonts are embedded or not.
    ///
    /// A font that was not collected gets the measured family, its fallback chain and a generic family.
    /// Without a render context only the family and sans-serif are written and nothing is embedded.
    /// 
    /// One instance serves one svg and is not thread-safe.
    /// </summary>
    internal class SvgFontPlan
    {
        private readonly RenderContext? _renderContext;
        private readonly bool _embedFonts;
        private readonly Dictionary<string, string> _stacks = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private WebFontSubsetPlanner? _planner;
        private WebFontPlan? _webFontPlan;

        internal SvgFontPlan(RenderContext? renderContext)
            : this(renderContext, false)
        {
        }

        /// <param name="renderContext">The context the render items were measured with, or null.</param>
        /// <param name="embedFonts">Whether fonts are embedded as subsets. Has no effect without a render context.</param>
        internal SvgFontPlan(RenderContext? renderContext, bool embedFonts)
        {
            _renderContext = renderContext;
            _embedFonts = embedFonts;
        }

        /// <summary>
        /// A plan without a render context: family and sans-serif only.
        /// </summary>
        internal static SvgFontPlan WithoutContext
        {
            get { return new SvgFontPlan(null); }
        }

        // -----------------------------------------------------------------------------------------
        // Collecting the text of the svg
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Starts collecting the text of a new svg and discards the plan of a previous one.
        /// </summary>
        internal void BeginCollect()
        {
            _stacks.Clear();
            _webFontPlan = null;
            _planner = _renderContext != null ? new WebFontSubsetPlanner(_renderContext.FontEngine, _embedFonts) : null;
        }

        /// <summary>
        /// Adds text measured with the font. A null or empty text only registers the font, so it
        /// gets a stack (for instance a paragraph's default font). Ignored without a render context.
        /// </summary>
        internal void AddText(IFontFormatBase? font, string? text)
        {
            if (_planner == null || font == null || string.IsNullOrEmpty(font.Family))
                return;
            _planner.AddText(font.Family, font.SubFamily, text);
        }

        /// <summary>
        /// Plans the stacks, and the subsets when fonts are embedded, for the collected text. Does
        /// nothing without a render context.
        /// </summary>
        internal void Build()
        {
            if (_planner == null)
                return;
            _webFontPlan = _planner.Build();
            _planner = null;
            _stacks.Clear();
        }

        // -----------------------------------------------------------------------------------------
        // Output
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Appends an @font-face rule for each embedded subset. Appends nothing when no font is embedded.
        /// The caller writes the style element around the rules.
        /// </summary>
        internal void AppendFontFaceRules(StringBuilder sb)
        {
            if (_webFontPlan == null || _webFontPlan.Faces.Count == 0)
                return;

            foreach (var face in _webFontPlan.Faces)
            {
                sb.Append("@font-face{font-family:");
                sb.Append(EscapeXml(QuoteFamilyName(face.FamilyName)));
                sb.Append(";src:url(data:font/ttf;base64,");
                sb.Append(Convert.ToBase64String(face.FontData));
                sb.Append(") format('truetype');font-weight:");
                sb.Append(face.FontWeight.ToString(CultureInfo.InvariantCulture));
                sb.Append(";font-style:");
                sb.Append(face.IsItalic ? "italic" : "normal");
                sb.Append(";}");
            }
        }

        /// <summary>
        /// Returns the value of the font-family attribute for the font, escaped for use inside
        /// a double-quoted xml attribute.
        /// </summary>
        internal string GetFontFamilyAttributeValue(IFontFormatBase? font)
        {
            if (font == null || string.IsNullOrEmpty(font.Family))
                return ToCss(GenericFontFamily.SansSerif);

            return GetFontFamilyAttributeValue(font.Family, font.SubFamily);
        }

        /// <summary>
        /// Returns the value of the font-family attribute for the family, escaped for use inside
        /// a double-quoted xml attribute.
        /// </summary>
        internal string GetFontFamilyAttributeValue(string family, FontSubFamily subFamily)
        {
            if (string.IsNullOrEmpty(family))
                return ToCss(GenericFontFamily.SansSerif);

            var key = family + "|" + subFamily;
            string stack;
            if (_stacks.TryGetValue(key, out stack))
                return stack;

            IReadOnlyList<string>? names = _webFontPlan != null ? _webFontPlan.GetFamilyNames(family, subFamily) : null;
            stack = names != null
                ? FormatStack(names, _webFontPlan!.GetGenericFamily(family, subFamily))
                : EscapeXml(BuildStack(family, subFamily));
            _stacks.Add(key, stack);
            return stack;
        }

        /// <summary>
        /// Splits the text where the font-family stack changes, which happens when measurement routed
        /// code points to a glyph fallback that is written by name (see
        /// <see cref="WebFontPlan.GetFamilyNames(string, FontSubFamily, int)"/>). Returns null when the
        /// whole text is drawn with the font's own stack, which is the normal case.
        /// </summary>
        internal List<SvgTextSegment>? GetFallbackSegments(IFontFormatBase? font, string? text)
        {
            if (_webFontPlan == null || font == null || string.IsNullOrEmpty(font.Family) || string.IsNullOrEmpty(text))
                return null;

            var fontStack = _webFontPlan.GetFamilyNames(font.Family, font.SubFamily);
            if (fontStack == null)
                return null;

            var segments = new List<SvgTextSegment>();
            IReadOnlyList<string>? current = null;
            int start = 0;
            bool split = false;

            for (int i = 0; i < text!.Length;)
            {
                int length = char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]) ? 2 : 1;
                int codePoint = length == 2 ? char.ConvertToUtf32(text[i], text[i + 1]) : text[i];

                // Marks, variation selectors and joiners stay with the character before them.
                IReadOnlyList<string> stack = i > 0 && IsAttachedToPrevious(text, i, codePoint)
                    ? current!
                    : _webFontPlan.GetFamilyNames(font.Family, font.SubFamily, codePoint);

                if (i == 0)
                {
                    current = stack;
                }
                else if (!ReferenceEquals(stack, current))
                {
                    segments.Add(CreateSegment(text.Substring(start, i - start), current!, fontStack, font));
                    start = i;
                    current = stack;
                    split = true;
                }
                if (!ReferenceEquals(stack, fontStack))
                    split = true;

                i += length;
            }

            if (!split)
                return null;

            segments.Add(CreateSegment(text.Substring(start), current!, fontStack, font));
            return segments;
        }

        /// <summary>
        /// The width of the text in points, measured as layout measured it. Null without a render
        /// context or when the font cannot be measured.
        /// </summary>
        internal double? MeasureWidthInPoints(IFontFormatBase? font, string? text)
        {
            if (_renderContext == null || font == null || string.IsNullOrEmpty(font.Family) || string.IsNullOrEmpty(text))
                return null;

            // The family is already substituted for the target, so the engine is asked directly.
            var shaper = _renderContext.FontEngine.GetMeasurementShaper(font.Family, font.SubFamily);
            if (shaper == null)
                return null;
            return shaper.ShapeLight(text).GetWidthInPoints(font.Size);
        }

        private SvgTextSegment CreateSegment(string text, IReadOnlyList<string> stack, IReadOnlyList<string> fontStack, IFontFormatBase font)
        {
            var fontFamily = ReferenceEquals(stack, fontStack)
                ? null
                : FormatStack(stack, _webFontPlan!.GetGenericFamily(font.Family, font.SubFamily));
            return new SvgTextSegment(text, fontFamily);
        }

        private static bool IsAttachedToPrevious(string text, int index, int codePoint)
        {
            switch (CharUnicodeInfo.GetUnicodeCategory(text, index))
            {
                case UnicodeCategory.NonSpacingMark:
                case UnicodeCategory.SpacingCombiningMark:
                case UnicodeCategory.EnclosingMark:
                case UnicodeCategory.Format:
                    return true;
                default:
                    // Emoji modifiers (skin tones) are symbols but belong to the emoji before them.
                    return codePoint >= 0x1F3FB && codePoint <= 0x1F3FF;
            }
        }

        private string BuildStack(string family, FontSubFamily subFamily)
        {
            var sb = new StringBuilder();
            sb.Append(QuoteFamilyName(family));

            var generic = GenericFontFamily.SansSerif;
            if (_renderContext != null)
            {
                var engine = _renderContext.FontEngine;
                foreach (var fallback in engine.GetFallbackChain(family))
                {
                    sb.Append(", ");
                    sb.Append(QuoteFamilyName(fallback));
                }
                generic = engine.GetGenericFamily(family, subFamily);
            }

            sb.Append(", ");
            sb.Append(ToCss(generic));
            return sb.ToString();
        }

        private static string FormatStack(IReadOnlyList<string> names, GenericFontFamily generic)
        {
            var sb = new StringBuilder();
            foreach (var name in names)
            {
                sb.Append(QuoteFamilyName(name));
                sb.Append(", ");
            }
            sb.Append(ToCss(generic));
            return EscapeXml(sb.ToString());
        }

        /// <summary>
        /// Quotes a family name as a CSS string. Quoting every name avoids the rules for unquoted
        /// names, which reject names such as "Noto Sans 3" or names that are generic keywords.
        /// </summary>
        internal static string QuoteFamilyName(string name)
        {
            return "'" + name.Replace("\\", "\\\\").Replace("'", "\\'") + "'";
        }

        internal static string ToCss(GenericFontFamily generic)
        {
            switch (generic)
            {
                case GenericFontFamily.Serif:
                    return "serif";
                case GenericFontFamily.Monospace:
                    return "monospace";
                default:
                    return "sans-serif";
            }
        }

        /// <summary>
        /// Escapes a value for use in xml text or inside a double-quoted xml attribute.
        /// </summary>
        private static string EscapeXml(string value)
        {
            return value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }
    }

    /// <summary>
    /// A part of a text run that is drawn with its own font-family stack.
    /// </summary>
    internal sealed class SvgTextSegment
    {
        internal SvgTextSegment(string text, string? fontFamily)
        {
            Text = text;
            FontFamily = fontFamily;
        }

        /// <summary>The text of the segment.</summary>
        internal string Text { get; private set; }

        /// <summary>
        /// The font-family attribute value, escaped for a double-quoted xml attribute, or null when the
        /// segment is drawn with the run's own stack.
        /// </summary>
        internal string? FontFamily { get; private set; }
    }
}
