/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/03/2026         EPPlus Software AB           Metrics fallback configuration
 *************************************************************************************************/
namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Controls whether text measurement may fall back to serialized font metrics when the
    /// requested font is not available as a font file.
    ///
    /// This affects measurement and line breaking only. Rendering and font embedding always
    /// require a real font file and are unaffected by this setting — PDF export in particular
    /// never uses serialized metrics, regardless of what is configured here.
    /// </summary>
    public enum MetricsFallbackMode
    {
        /// <summary>
        /// Never use serialized metrics. Text in a font that cannot be resolved is measured with
        /// the embedded last-resort font, which is condensed and therefore a poor width match for
        /// most proportional fonts.
        /// </summary>
        Disabled = 0,

        /// <summary>
        /// Default. Use serialized metrics when font resolution has fallen all the way through to
        /// the embedded last-resort font and the requested family has metrics available.
        ///
        /// The user-configured and built-in fallback chains still take precedence: a real,
        /// metric-compatible substitute font is a better measurement source than quantized
        /// metrics, since it also carries kerning and OpenType layout tables.
        /// </summary>
        WhenFontMissing = 1,

        /// <summary>
        /// Always measure from serialized metrics, ignoring installed fonts entirely.
        ///
        /// Measurement then does not depend on which fonts the machine has, so line breaking is
        /// reproducible across machines and testable without a font fixture. The cost is that
        /// kerning and OpenType substitutions are not applied and widths are quantized, so
        /// measurements differ slightly from what a real font would give.
        /// </summary>
        Always = 2
    }
}