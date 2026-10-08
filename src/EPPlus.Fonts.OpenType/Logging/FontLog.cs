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
using System;

namespace EPPlus.Fonts.OpenType.Logging
{
    /// <summary>
    /// Helpers shared by everything that raises font log events. All calls into a logger go
    /// through here, so an exception in a user-supplied logger never reaches the caller.
    /// </summary>
    internal static class FontLog
    {
        internal static bool IsEnabled(IFontLogger logger, FontLogSeverity severity)
        {
            if (logger == null)
                return false;
            try
            {
                return logger.IsEnabled(severity);
            }
            catch
            {
                return false;
            }
        }

        internal static void Write(IFontLogger logger, FontLogEvent logEvent)
        {
            if (logger == null || logEvent == null)
                return;
            try
            {
                logger.Log(logEvent);
            }
            catch
            {
                // Logging must never break font resolution or rendering.
            }
        }

        internal static void Write(
            IFontLogger logger,
            FontLogSeverity severity,
            FontLogEventType type,
            string message,
            string requestedFont,
            string resolvedFont)
        {
            Write(logger, new FontLogEvent
            {
                Type = type,
                Severity = severity,
                Message = message,
                RequestedFont = requestedFont,
                ResolvedFont = resolvedFont
            });
        }

        /// <summary>Family and subfamily of a font, for messages.</summary>
        internal static string Describe(OpenTypeFont font)
        {
            if (font == null)
                return "(none)";
            try
            {
                return font.GetEnglishFontFamilyName() + " " + font.SubFamily;
            }
            catch
            {
                return "(unknown)";
            }
        }

        /// <summary>Family name of a font, or null when it cannot be read.</summary>
        internal static string FamilyOf(OpenTypeFont font)
        {
            if (font == null)
                return null;
            try
            {
                return font.GetEnglishFontFamilyName();
            }
            catch
            {
                return null;
            }
        }

        internal static string JoinNames(string[] names)
        {
            if (names == null || names.Length == 0)
                return "none";
            return string.Join(", ", names);
        }

        internal static string FormatCodePoint(uint codePoint)
        {
            return "U+" + codePoint.ToString("X4");
        }
    }
}