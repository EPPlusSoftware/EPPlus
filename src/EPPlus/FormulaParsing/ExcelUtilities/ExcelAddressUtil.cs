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
    /// Utilites tp verify addresses and reöated tokens
    /// </summary>
    public static class ExcelAddressUtil
    {
        static char[] SheetNameInvalidChars = new char[] { '?', ':', '*', '/', '\\' };
        /// <summary>
        /// Ensure address and sheet has valid names 
        /// </summary>
        /// <param name="token"></param>
        /// <returns>Wether or not the address is valid</returns>
        public static bool IsValidAddress(string token)
        {
            int ix;
            if (token[0] == '\'')
            {
                ix = token.LastIndexOf('\'');
                if (ix > 0 && ix < token.Length - 1 && token[ix + 1] == '!')
                {
                    if (token.IndexOfAny(SheetNameInvalidChars, 1, ix - 1) > 0)
                    {
                        return false;
                    }
                    token = token.Substring(ix + 2);
                }
                else
                {
                    return false;
                }
            }
            else if ((ix = token.IndexOf('!')) > 1)
            {
                if (token.IndexOfAny(SheetNameInvalidChars, 0, token.IndexOf('!')) > 0)
                {
                    return false;
                }
                token = token.Substring(token.IndexOf('!') + 1);
            }
            return ExcelCellBase.IsValidAddress(token);
        }
        readonly static char[] NameInvalidChars = new char[] { '!', '@', '#', '$', '£', '%', '&', '/', '(', ')', '[', ']', '{', '}', '<', '>', '=', '+', '*', '-', '~', '^', ':', ';', '|', ',', ' ', '\t', '\r', '\n' };
        /// <summary>
        /// Returns true if a defined name is valid
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static bool IsValidName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }
            var fc = name[0];
            if (!(char.IsLetter(fc) || fc == '_' || (fc == '\\' && name.Length != 2)))
            {
                return false;
            }

            if (name.IndexOfAny(NameInvalidChars, 1) > 0)
            {
                return false;
            }

            if(ExcelCellBase.IsValidAddress(name))
            {
                return false;
            }

            //TODO:Add check for functionnames.
            return true;
        }
        /// <summary>
        /// Ensures valid name by removing invalid chars and replacing them with '_'
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static string GetValidName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return name;
            }

            var fc = name[0];
            if (!(char.IsLetter(fc) || fc == '_' || (fc == '\\' && name.Length > 2)))
            {
                name = "_" + name.Substring(1);
            }

            name = NameInvalidChars.Aggregate(name, (c1, c2) => c1.Replace(c2, '_'));
            return name;
        }
    }
}
