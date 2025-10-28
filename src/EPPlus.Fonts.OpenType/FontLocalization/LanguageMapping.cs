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

namespace EPPlus.Fonts.OpenType.FontLocalization
{
    public class LanguageMapping
    {
        public int code { get; set; }

        public Languages Language { get; set; }

        internal static LanguageMapping Create(int code, Languages language)
        {
            return new LanguageMapping
            {
                code = code,
                Language = language
            };
        }

        public override string ToString()
        {
            return Language.ToString();
        }
    }
}
