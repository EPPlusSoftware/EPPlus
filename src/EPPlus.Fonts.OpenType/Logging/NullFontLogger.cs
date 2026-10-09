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

namespace EPPlus.Fonts.OpenType.Logging
{
    /// <summary>
    /// Used when no logger is configured. Never enabled, so no event is ever built.
    /// </summary>
    internal sealed class NullFontLogger : IFontLogger
    {
        internal static readonly NullFontLogger Instance = new NullFontLogger();

        private NullFontLogger()
        {
        }

        public bool IsEnabled(FontLogSeverity severity)
        {
            return false;
        }

        public void Log(FontLogEvent logEvent)
        {
        }
    }
}