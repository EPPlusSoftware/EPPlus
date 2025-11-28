/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using System;

namespace EPPlus.Fonts.OpenType.FontValidation
{
    internal static class LongDateTime
    {
        // OpenType epoch: 1904-01-01 00:00:00 UTC
        private static readonly DateTime EpochUtc = new DateTime(1904, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        /// <summary>
        /// Converts OpenType LONGDATETIME (seconds since 1904-01-01 UTC) to DateTime (UTC).
        /// Returns null if value is obviously invalid (e.g., less than MinValue or overflow).
        /// </summary>
        public static DateTime? ToDateTimeUtc(long secondsSinceEpoch)
        {
            try
            {
                // Guard against clearly invalid negatives far before epoch (spec says seconds since epoch; negatives are unusual)
                // We allow 0 (exact epoch) and positive values; small negatives can be treated as invalid.
                if (secondsSinceEpoch < 0)
                    return null;

                // AddSeconds internally checks for overflow; we use checked for safety.
                DateTime dt = EpochUtc.AddSeconds(secondsSinceEpoch);
                return dt;
            }
            catch
            {
                return null;
            }
        }
    }
}
