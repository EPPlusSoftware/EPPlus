using System.Net;
using System.Text.RegularExpressions;

namespace EPPlus.DrawingRenderer.Tests
{
    /// <summary>
    /// Reads properties from the style attribute of svg elements in tests. Font properties are written
    /// in the style attribute rather than as presentation attributes, so that an inline style wins over
    /// the css of a host page.
    /// </summary>
    internal static class SvgStyle
    {
        private static readonly Regex StyleAttribute = new Regex("\\bstyle=\"([^\"]*)\"");

        /// <summary>
        /// The value of a property in the style attribute of an element, decoded from xml, or null when
        /// the element has no style attribute or the style has no such property.
        /// </summary>
        /// <param name="element">The start tag of the element, or its attributes.</param>
        /// <param name="property">The css property, e.g. "font-family".</param>
        internal static string? Property(string element, string property)
        {
            var style = StyleAttribute.Match(element);
            if (!style.Success)
                return null;

            var declarations = WebUtility.HtmlDecode(style.Groups[1].Value);
            var match = Regex.Match(declarations, "(?:^|;)\\s*" + Regex.Escape(property) + "\\s*:\\s*([^;]*)");
            return match.Success ? match.Groups[1].Value.Trim() : null;
        }
    }
}