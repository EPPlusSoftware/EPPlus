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
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.FontResolver
{
    public class FontResolutionConfig
    {
        private readonly Dictionary<string, string[]> _fallbacks =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);

        public FontResolutionConfig AddFallback(string primaryFont, params string[] fallbacks)
        {
            if (string.IsNullOrEmpty(primaryFont))
                throw new ArgumentNullException("primaryFont");
            if (fallbacks == null || fallbacks.Length == 0)
                throw new ArgumentException("At least one fallback must be specified", "fallbacks");

            _fallbacks[primaryFont] = fallbacks;
            return this;
        }

        internal string[] GetFallbacks(string fontName)
        {
            string[] result;
            return _fallbacks.TryGetValue(fontName, out result) ? result : null;
        }
    }
}
