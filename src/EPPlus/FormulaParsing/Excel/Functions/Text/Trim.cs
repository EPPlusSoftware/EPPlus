﻿/*************************************************************************************************
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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "4",
        Description = "Removes duplicate spaces, and spaces at the start and end of a text string",
        SupportsArrays = true)]
    internal class Trim : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var str = ArgToString(arguments, 0);
            if(!string.IsNullOrEmpty(str))
            {
                str = str.Trim();
                str = Regex.Replace(str, "[ ]+", " ");
            }
            return CreateResult(str, DataType.String);
        }
    }
}
