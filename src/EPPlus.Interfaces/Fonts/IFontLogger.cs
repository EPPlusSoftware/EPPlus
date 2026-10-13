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
    /// Receives diagnostic events describing font and glyph selection.
    /// Attach an implementation via <see cref="IEpplusFontConfiguration.Logger"/>.
    /// </summary>
    /// <remarks>
    /// Implementations must be thread-safe: events are raised on whichever thread triggers font
    /// resolution or text shaping. Exceptions thrown by an implementation are swallowed, so a
    /// faulty logger can never break rendering.
    /// </remarks>
    public interface IFontLogger
    {
        /// <summary>
        /// Called before an event is built. Return false to skip events of this severity, which
        /// avoids the cost of formatting messages that would be discarded.
        /// </summary>
        bool IsEnabled(FontLogSeverity severity);

        /// <summary>
        /// Called for each event whose severity <see cref="IsEnabled"/> accepted.
        /// </summary>
        void Log(FontLogEvent logEvent);
    }
}