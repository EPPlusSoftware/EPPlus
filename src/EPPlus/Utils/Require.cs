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

namespace OfficeOpenXml.Utils
{
    /// <summary>
    /// Utility for validation in functions.
    /// </summary>
    public static class Require
    {
        /// <summary>
        /// Represent an argument to the function where the validation is implemented.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="argument">The argument to validate</param>
        /// <returns></returns>
        public static IArgument<T> Argument<T>(T argument)
        {
            return new Argument<T>(argument);
        }


    }
}
