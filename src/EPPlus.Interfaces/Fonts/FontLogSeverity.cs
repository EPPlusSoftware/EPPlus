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
namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Severity of a font diagnostic event.
    /// </summary>
    public enum FontLogSeverity
    {
        /// <summary>
        /// Detailed decisions, such as every exact match or every glyph routed to a fallback font.
        /// Can be verbose.
        /// </summary>
        Debug = 0,
        /// <summary>
        /// Decisions worth following, such as a font substituted by a fallback chain.
        /// </summary>
        Information = 1,
        /// <summary>
        /// The output may differ from what was requested, such as the last-resort font being used
        /// or a glyph being missing from every candidate font.
        /// </summary>
        Warning = 2,
        /// <summary>
        /// An operation was refused or failed, such as a font that may not be embedded.
        /// </summary>
        Error = 3
    }
}