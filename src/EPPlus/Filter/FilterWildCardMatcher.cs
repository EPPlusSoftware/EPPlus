/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Filter
{
    internal  static class FilterWildCardMatcher
    {
        internal static bool Match(string value, string pattern)
        {
            var tokens = SplitInTokens(pattern);
            if (tokens.Count == 1 && tokens[0]!="*" && tokens[0] != "?")
            {
                return value.Equals(tokens[0], StringComparison.CurrentCultureIgnoreCase);
            }
            return MatchTokenList(value, tokens, 0, 0);
        }

        private static bool MatchTokenList(string value, List<string> tokens, int stringPos, int tokenPos)
        {
            bool match = true;
            bool isWC=false;
            for (int i=tokenPos;i<tokens.Count;i++)
            {
                if (tokens[i] == "*")
                {
                    isWC = true;
                }
                else if (tokens[i] == "?")
                {
                    stringPos++;
                }
                else
                {
                    if (isWC)
                    {
                        return MatchWildCards(value, tokens, stringPos, i);
                    }
                    else if (stringPos + tokens[i].Length <= value.Length)
                    {
                        match = value.Substring(stringPos, tokens[i].Length).Equals(tokens[i], StringComparison.CurrentCultureIgnoreCase);
                        stringPos += tokens[i].Length;
                        isWC = false;
                    }
                    else
                    {
                        return false;
                    }
                }
                if (match == false) return false;
            }
            if (isWC)
            {
                return true;
            }
            else
            {
                return stringPos == value.Length;
            }
        }
        private static bool MatchWildCards(string value, List<string> tokens, int stringPos, int tokenPos)
        {
            var anyChars = 0;
            while (tokens[tokenPos]=="*" || tokens[tokenPos] == "?")
            {
                if(tokens[tokenPos] == "?")
                {
                    anyChars++;
                }
                tokenPos++;
                if (tokenPos == tokens.Count) return (value.Length-stringPos)>anyChars;
            }
            stringPos += anyChars;
            if (stringPos > value.Length) return false;

            int foundPos = value.IndexOf(tokens[tokenPos], stringPos, StringComparison.CurrentCultureIgnoreCase);
            while (foundPos>=0)
            {
                bool match;
                if (tokenPos + 1 >= tokens.Count)
                {
                    match = foundPos + tokens[tokenPos].Length == value.Length;
                    if (match) return true;
                }
                else
                {
                    match = MatchTokenList(value, tokens, foundPos + tokens[tokenPos].Length, tokenPos + 1);
                    if (match) return true;
                }
                foundPos = value.IndexOf(tokens[tokenPos], foundPos+1, StringComparison.CurrentCultureIgnoreCase);
            }
            return false;
        }

        private static void FindMatchingWildCards(string v1, string value, List<string> list, List<string> tokens, int v2, int stringPos, int v3, int tokenPos)
        {
            throw new NotImplementedException();
        }

        private static List<string> SplitInTokens(string filter)
        {
            var ret = new List<string>();
            var start = 0;
            for (int i = 0; i < filter.Length; i++)
            {
                if (filter[i] == '*' ||
                   filter[i] == '?')
                {
                    if (start < i)
                    {
                        ret.Add(filter.Substring(start, i - start).Replace("**","*").Replace("??", "?"));
                    }
                    ret.Add(filter[i].ToString());
                    start = i + 1;
                }
            }
            if (start < filter.Length)
            {
                ret.Add(filter.Substring(start, filter.Length - start).Replace("**", "*").Replace("??", "?"));
            }
            return ret;
        }
    }
}
