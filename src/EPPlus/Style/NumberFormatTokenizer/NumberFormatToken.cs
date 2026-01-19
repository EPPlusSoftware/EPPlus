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
    internal class NumberFormatToken
    {
        public NumberFormatTokenType TokenType { get; set; }
        public string Value { get; set; }
        public NumberFormatToken(NumberFormatTokenType tokenType, string value)
        {
            TokenType = tokenType;
            Value = value;
        }
    }
}