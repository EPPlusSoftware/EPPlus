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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Returns an integer representing the day of the week for a supplied date",
        SupportsArrays = true)]
    internal class Weekday : ExcelFunction
    {
        internal override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var serialNumber = ArgToDecimal(arguments, 0);
            var returnType = arguments.Count() > 1 ? ArgToInt(arguments, 1) : 1;
            return CreateResult(CalculateDayOfWeek(System.DateTime.FromOADate(serialNumber), returnType), DataType.Integer);
        }

        private static List<int> _oneBasedStartOnSunday = new List<int> { 1, 2, 3, 4, 5, 6, 7 };
        private static List<int> _oneBasedStartOnMonday = new List<int> { 7, 1, 2, 3, 4, 5, 6 };
        private static List<int> _zeroBasedStartOnSunday = new List<int> { 6, 0, 1, 2, 3, 4, 5 };

        private int CalculateDayOfWeek(System.DateTime dateTime, int returnType)
        {
            var dayIx = (int)dateTime.DayOfWeek;
            switch (returnType)
            {
                case 1:
                    return _oneBasedStartOnSunday[dayIx];
                case 2:
                case 11:
                    return _oneBasedStartOnMonday[dayIx];
                case 3:
                    return _zeroBasedStartOnSunday[dayIx];
                default:
                    throw new ExcelErrorValueException(eErrorType.Num);
            }
        }
    }
}
