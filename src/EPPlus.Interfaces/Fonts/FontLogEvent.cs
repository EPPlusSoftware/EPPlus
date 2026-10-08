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
using System;

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// A diagnostic event describing a font or glyph selection decision.
    /// <see cref="Message"/> is a ready-made line of text; the other properties carry the same
    /// information in structured form for loggers that filter or assert on it.
    /// </summary>
    public sealed class FontLogEvent
    {
        /// <summary>The kind of decision.</summary>
        public FontLogEventType Type { get; set; }

        /// <summary>The severity of the event.</summary>
        public FontLogSeverity Severity { get; set; }

        /// <summary>A human-readable description.</summary>
        public string Message { get; set; }

        /// <summary>The font that was asked for, when applicable. For glyph events, the primary font.</summary>
        public string RequestedFont { get; set; }

        /// <summary>The font that was chosen, when applicable.</summary>
        public string ResolvedFont { get; set; }

        /// <summary>The Unicode script involved, for glyph and script events.</summary>
        public UnicodeScript? Script { get; set; }

        /// <summary>The code point involved, for glyph events.</summary>
        public uint? CodePoint { get; set; }

        /// <summary>The exception that caused the event, when there is one.</summary>
        public Exception Exception { get; set; }

        /// <inheritdoc/>
        public override string ToString()
        {
            return Message;
        }
    }
}