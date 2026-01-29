/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
namespace OfficeOpenXml.Style
{
    public enum NumberFormatTokenType
    {
        NumberFormat,
        DateTimeFormat,
        Semicolon,
        String,
        StringContent,
        LanguageCalenderString,
        CharacterWidth,
        Text, //A literal text token, not inside a string.
        Color,
        Condition,
        ElapsedTime,
        /// <summary>
        /// The parsed token represents an opening bracket ('[')
        /// </summary>
        OpeningBracket,
        /// <summary>
        /// The parsed token represents a closing bracket (']')
        /// </summary>
        ClosingBracket,
    }
}