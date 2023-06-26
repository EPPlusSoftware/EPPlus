/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/16/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    /// <summary>
    /// Compares values against wildcard strings
    /// </summary>
    internal class WildCardValueMatcher2 : ValueMatcher
    {
        /// <summary>
        /// Compares two strings
        /// </summary>
        /// <param name="searchedValue">The searched value, might contain wildcard characters</param>
        /// <param name="candidate">The candidate to compare</param>
        /// <returns>0 if match, otherwise -1 or 1</returns>
        protected override int CompareStringToString(string searchedValue, string candidate)
        {
            if (searchedValue.Contains("*") || searchedValue.Contains("?"))
            {
                if (IsMatch(searchedValue, candidate))
                {
                    return 0;
                }
            }
            return base.CompareStringToString(candidate, searchedValue);
        }

        private bool IsMatch(string searchedValue, string candidate)
        {
            int svIx = 0;
            int cIx = 0;
            var pattern = searchedValue.ToUpperInvariant();
            var cand = candidate.ToUpperInvariant();
            bool escapeNextWildCard = false;
            do
            {
                var sv = pattern[svIx];
                if(!escapeNextWildCard && sv == '*' && svIx == pattern.Length - 1)
                {
                    // if the last char of the searched value
                    // is an asterix and we have made it to
                    // here it's a match.
                    return true;
                }
                else if (
                    sv == '~' 
                    && svIx < pattern.Length - 1 
                    && (pattern[svIx + 1] == '*' || pattern[svIx + 1] == '?')
                    )
                {
                    // current char is an escape char
                    // and the next char is a wildcard
                    svIx++; 
                    escapeNextWildCard = true;
                    continue;
                }
                if (sv == '*' && !escapeNextWildCard)
                {
                    // if multiple *'s just ignore them
                    if(svIx < pattern.Length - 1)
                    {
                        var tmpIx = svIx + 1;
                        while (tmpIx < pattern.Length && pattern[tmpIx] == '*')
                        {
                            tmpIx++;
                        }
                        svIx = tmpIx;
                    }
                    var cont = false;
                    var svPart = new StringBuilder();
                    var svC = pattern[svIx];
                    do
                    {
                        if(svC == '~')
                        {
                            if(svIx < pattern.Length -1)
                            {
                                var escCand = pattern.Substring(svIx, 2);
                                if(escCand == "~*" || escCand == "~?")
                                {
                                    escapeNextWildCard = true;
                                    svIx++;
                                    svC = pattern[svIx];
                                    cIx = cand.IndexOf(svC);
                                    //cont = true;
                                    //continue;
                                }
                            }
                        }
                        svPart.Append(svC);
                        svIx++;
                        if (svIx < pattern.Length)
                            svC = pattern[svIx];
                    }
                    while ((svC != '*' && svC != '?' && svC != '~') && svIx < pattern.Length);
                    if (cont) continue;
                    var part = svPart.ToString();
                    if (cand.EndsWith(part) && svIx == pattern.Length) return true;
                    cIx = cand.IndexOf(part);
                    if (cIx < 0) return false;
                    cIx += part.Length;
                    
                }
                else if(svIx < pattern.Length -1 && sv == '~')
                {
                    var next = pattern[svIx + 1];
                    if(next == '*' || next == '?')
                    {
                        escapeNextWildCard= true;
                        svIx++;
                    }
                    else if (cand[cIx] != '~')
                    { 
                        return false;
                    }
                }
                else if (sv == '?' && !escapeNextWildCard)
                {
                    if (cIx > cand.Length - 1) return false;
                    cIx++;
                    svIx++;
                }
                else if (cIx < cand.Length && sv == cand[cIx])
                {
                    cIx++;
                    svIx++;
                    escapeNextWildCard = false;
                }
                else
                {
                    return false;
                }
            }
            while (svIx < pattern.Length);
            if(cIx < cand.Length - 1 && pattern.Last() != '*')
            {
                return false;
            }
            return true;
        }
    }
}
