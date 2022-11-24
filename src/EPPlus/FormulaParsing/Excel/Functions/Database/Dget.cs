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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Database,
        EPPlusVersion = "4",
        Description = "Returns a single value from a field of a list or database, that satisfy specified conditions")]
    internal class Dget : DatabaseFunction
    {

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var rowMatcher = new RowMatcher(context);
            var dbAddress = arguments.ElementAt(0).ValueAsRangeInfo.Address.ToString();
            var field = ArgToString(arguments, 1).ToLower(CultureInfo.InvariantCulture);
            var criteriaRange = arguments.ElementAt(2).ValueAsRangeInfo.Address.ToString();

            var db = new ExcelDatabase(context.ExcelDataProvider, dbAddress);
            var criteria = new ExcelDatabaseCriteria(context.ExcelDataProvider, criteriaRange);

            var nHits = 0;
            object retVal = null;
            while (db.HasMoreRows)
            {
                var dataRow = db.Read();
                if (!rowMatcher.IsMatch(dataRow, criteria)) continue;
                if(++nHits > 1) return CreateResult(ExcelErrorValue.Values.Num, DataType.ExcelError);
                retVal = dataRow[field];
            }
            return CompileResultFactory.Create(retVal);
        }
    }
}
