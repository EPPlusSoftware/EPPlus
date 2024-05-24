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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Int argument parser
    /// </summary>
    internal class IntArgumentParser : ArgumentParser
    {
        /// <summary>
        /// Parse object to int
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override object Parse(object obj)
        {
            Require.That(obj).Named("argument").IsNotNull();
            return Parse(obj, RoundingMethod.Convert);
        }
        /// <summary>
        /// Parse object to int roundingMethod
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="roundingMethod"></param>
        /// <returns></returns>
        public override object Parse(object obj, RoundingMethod roundingMethod)
        {
            return Utils.ConvertUtil.ParseInt(obj, roundingMethod);
        }
    }
}
