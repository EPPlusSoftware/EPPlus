/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/08/2026         EPPlus Software AB           Font logging
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;
using System.IO;

namespace EPPlus.Fonts.OpenType.Logging
{
    /// <summary>
    /// Creates ready-made <see cref="IFontLogger"/> implementations.
    /// </summary>
    /// <example>
    /// workbook.ConfigureFonts(c =>
    /// {
    ///     c.Logger = FontLoggerFactory.CreateTextFileLogger(new FileInfo(@"c:\fontlog.txt"));
    /// });
    /// </example>
    public static class FontLoggerFactory
    {
        /// <summary>
        /// Creates a logger that writes one line per event to <paramref name="logfile"/>, from
        /// <see cref="FontLogSeverity.Information"/> and up. An existing file is overwritten.
        /// </summary>
        /// <param name="logfile">The file to write to.</param>
        public static IFontLogger CreateTextFileLogger(FileInfo logfile)
        {
            return CreateTextFileLogger(logfile, FontLogSeverity.Information);
        }

        /// <summary>
        /// Creates a logger that writes one line per event to <paramref name="logfile"/>.
        /// An existing file is overwritten.
        /// </summary>
        /// <param name="logfile">The file to write to.</param>
        /// <param name="minimumSeverity">Events below this severity are not written.
        /// Use <see cref="FontLogSeverity.Debug"/> to follow every exact match and every glyph routed to a fallback.</param>
        public static IFontLogger CreateTextFileLogger(FileInfo logfile, FontLogSeverity minimumSeverity)
        {
            return new TextFileFontLogger(logfile, minimumSeverity);
        }
    }
}