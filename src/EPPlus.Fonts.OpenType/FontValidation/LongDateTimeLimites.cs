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

    internal static class LongDateTimeLimits
    {
        // Epoch
        private static readonly DateTime EpochUtc = new DateTime(1904, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        // Precompute boundaries in seconds since epoch
        public static readonly long MinSeconds = 0L; // 1904-01-01
        public static long MaxSeconds // now + 10 years
        {
            get
            {
                DateTime max = DateTime.UtcNow.AddYears(10);
                TimeSpan span = max - EpochUtc;
                return (long)span.TotalSeconds;
            }
        }
    }

}
