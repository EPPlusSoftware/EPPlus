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
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Style
{
    internal class NumberFormatTokenizer 
    {
        internal static IList<NumberFormatToken> Tokenize(string input)
        {
            var tokens = new List<NumberFormatToken>();
            bool isInString = false;
            var value = new StringBuilder();
            for (int i=0;i<input.Length;i++)
            {
                var c = input[i];
                if(isInString)
                {
                    if(c=='"')
                    {
                        if(i>0 && input[i - 1]=='\\')
                        {
                            //Escaped quote
                            value.Append(c);
                            i++;
                            continue;
                        }
                        isInString = false;
                        tokens.Add(new NumberFormatToken(NumberFormatTokenType.StringContent, value.ToString()));
                        value = new StringBuilder();
                        tokens.Add(new NumberFormatToken(NumberFormatTokenType.String, c.ToString()));
                    }
                    else
                    {
                        value.Append(c);
                    }

                    continue;
                }                
                switch(c)
                {
                    case '"':
                        tokens.Add(new NumberFormatToken(NumberFormatTokenType.Text, value.ToString()));
                        if (value.Length > 0)
                        {
                            value = new StringBuilder();
                        }
                        isInString = true;
                        continue;
                    case '[':
                        tokens.Add(new NumberFormatToken(NumberFormatTokenType.OpeningBracket, c.ToString()));
                        continue;
                    case ']':
                        HandleBracketContent(ref value, tokens);
                        tokens.Add(new NumberFormatToken(NumberFormatTokenType.ClosingBracket, c.ToString()));
                        break;
                    case ';':
                        var v = value.ToString();
                        tokens.Add(new NumberFormatToken(GetTokenTypeFromValue(v), v));
                        tokens.Add(new NumberFormatToken(NumberFormatTokenType.Semicolon, c.ToString()));
                        break;
                    case '_':
                        if (i + 1 < input.Length)
                        {
                            tokens.Add(new NumberFormatToken(NumberFormatTokenType.CharacterWidth, input[i + 1].ToString()));
                            i++;
                        }
                        else
                        {
                            throw (new InvalidOperationException("Invalid number format string. Ending with a _ is not valid"));
                        }
                        break;
                    default:
                        value.Append(c);
                        break;
                }
            }
            return tokens;
        }

        private static NumberFormatTokenType GetTokenTypeFromValue(string v)
        {
            if(v.IndexOfAny(new char[] { 'y', 'd','m', 'h', 's' }, 0)>=0 && v.IndexOfAny(new char[] { '0', '#', '?'})<0)
            {
                return NumberFormatTokenType.DateTimeFormat;
            }
            else
            {
                return NumberFormatTokenType.NumberFormat;
            }
        }

        internal static HashSet<string> FormatColors = new HashSet<string>(
            new string[]{
                "black",
                "blue",
                "cyan",
                "green",
                "magenta",
                "red",
                "white",
                "yellow"
        }, StringComparer.OrdinalIgnoreCase);

        private static void HandleBracketContent(ref StringBuilder value, List<NumberFormatToken> tokens)
        {
            var v = value.ToString();
            if (FormatColors.Contains(v))
            {
                tokens.Add(new NumberFormatToken(NumberFormatTokenType.Color, v));
            }
            else if (v.StartsWith("=") ||
               v.StartsWith(">") ||
               v.StartsWith("<"))
            {
                tokens.Add(new NumberFormatToken(NumberFormatTokenType.Condition, v));
            }
            else if (v.StartsWith("$") || v.StartsWith("0xf"))
            {
                tokens.Add(new NumberFormatToken(NumberFormatTokenType.LanguageCalenderString, v));
            }
            else if (v.StartsWith("h") || v.StartsWith("m") || v.StartsWith("s"))
            {
                tokens.Add(new NumberFormatToken(NumberFormatTokenType.ElapsedTime, v));
            }
            else
            {
                throw (new InvalidOperationException("Invalid format inside brackaets[]"));
            }
        }
    }
}