﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/15/2024         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Database,
        EPPlusVersion = "7",
        Description = "Multiplies the values in a field of a list or database, that satisfy specified conditions")]
    internal class Dproduct : DatabaseFunction
    {
        public override int ArgumentMinLength => 3;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var values = GetMatchingValues(arguments, context);
            if (!values.Any()) return CreateResult(0d, DataType.Integer);
            return CreateResult(values.Aggregate(1d, (accumulated, val) => accumulated * val), DataType.Integer);
        }
        /// <summary>
        /// If the function is allowed in a pivot table calculated field
        /// </summary>
        public override bool IsAllowedInCalculatedPivotTableField => false;
    }
}