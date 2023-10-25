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
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    /// <summary>
    /// Handles wildcard matches in functions.
    /// </summary>
    public class WildCardValueMatcher : ValueMatcher
    {
        /// <summary>
        /// Compares string to string allowing wildcards
        /// </summary>
        /// <param name="searchedValue">The value to search for</param>
        /// <param name="candidate">The candidate object</param>
        /// <returns>The compare result</returns>
        protected override int CompareStringToString(string searchedValue, string candidate)
        {
            if (searchedValue.Contains("*") || searchedValue.Contains("?"))
            {
                var regexPattern = BuildRegex(searchedValue, candidate);
                if (Regex.IsMatch(candidate, regexPattern))
                {
                    return 0;
                }
            }
            return base.CompareStringToString(candidate, searchedValue);
        }

        private string BuildRegex(string searchedValue, string candidate)
        {
            var result = new StringBuilder();
            var regexPattern = Regex.Escape(searchedValue);
            regexPattern = regexPattern.Replace("\\*", "*");
            regexPattern = regexPattern.Replace("\\?", "?");
            regexPattern = string.Format("^{0}$", regexPattern);
            var lastIsTilde = false;
            foreach(var ch in regexPattern)
            {
                if(ch == '~')
                {
                    lastIsTilde = true;
                    continue;
                }
                if(ch == '*')
                {
                    if(lastIsTilde)
                    {
                        result.Append("\\*");
                    }
                    else
                    {
                        result.Append(".*");
                    }
                }
                else if(ch == '?')
                {
                    if (lastIsTilde)
                    {
                        result.Append("\\?");
                    }
                    else
                    {
                        result.Append('.');
                    }
                }
                else if(lastIsTilde)
                {
                    result.Append("~" + ch);
                }
                else
                {
                    result.Append(ch);
                }
                lastIsTilde = false;
            }
            return result.ToString();
        }
    }
}
