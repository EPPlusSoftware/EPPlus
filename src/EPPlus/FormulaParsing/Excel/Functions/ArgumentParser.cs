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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Argument parser base abstract class
    /// </summary>
    internal abstract class ArgumentParser
    {
        /// <summary>
        /// Parse object argument
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public abstract object Parse(object obj);
        /// <summary>
        /// Parse object argument and round it
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="roundingMethod"></param>
        /// <returns></returns>
        public abstract object Parse(object obj, RoundingMethod roundingMethod);
    }
}
