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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml
{

    internal static class ErrorValues
    {
        public static ExcelErrorValue ValueError = ExcelErrorValue.Create(eErrorType.Value);
        public static ExcelErrorValue NameError = ExcelErrorValue.Create(eErrorType.Name);
        public static ExcelErrorValue NAError = ExcelErrorValue.Create(eErrorType.NA);
        public static ExcelErrorValue NumError = ExcelErrorValue.Create(eErrorType.Num);
        public static ExcelErrorValue NullError = ExcelErrorValue.Create(eErrorType.Null);
        public static ExcelErrorValue Div0Error = ExcelErrorValue.Create(eErrorType.Div0);
        public static ExcelErrorValue RefError = ExcelErrorValue.Create(eErrorType.Ref);
        public static ExcelErrorValue CalcError = ExcelErrorValue.Create(eErrorType.Calc);
    }
}
