/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    internal static class PaddingHelper
    {
        public static string EnsureLength(string input, int length, string padWith = "")
        {
            if (input == null) input = string.Empty;
            if (input.Length < length && !string.IsNullOrEmpty(padWith))
            {
                while (input.Length < length)
                {
                    input = padWith + input;
                }
            }
            else if (input.Length > length)
            {
                input = input.Substring(input.Length - length);
            }
            return input;
        }

        public static string EnsureMinLength(string input, int length)
        {
            if (input.Length > length)
            {
                input = input.Substring(input.Length - length);
            }
            return input;
        }
    }
}
