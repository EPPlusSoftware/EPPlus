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
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Boolean argument parser
    /// </summary>
    internal class BoolArgumentParser : ArgumentParser
    {
        /// <summary>
        /// Parse object to bool
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override object Parse(object obj)
        {
            return ConvertUtil.GetValueBool(obj);
        }
        /// <summary>
        /// Parse object to bool with rounding method
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="roundingMethod"></param>
        /// <returns></returns>
        public override object Parse(object obj, RoundingMethod roundingMethod)
        {
            return Parse(obj);
        }
    }
}
