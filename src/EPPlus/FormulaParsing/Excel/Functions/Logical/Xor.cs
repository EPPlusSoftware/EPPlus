/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    [FunctionMetadata(
            Category = ExcelFunctionCategory.Logical,
            EPPlusVersion = "5.5",
            Description = "Returns a logical Exclusive Or of all arguments",
            IntroducedInExcelVersion = "2013")]
    internal class Xor : ExcelFunction
    {
        public Xor()
            : this(new DoubleEnumerableArgConverter())
        {

        }

        public Xor(DoubleEnumerableArgConverter converter)
        {
            _converter = converter;
        }

        private readonly DoubleEnumerableArgConverter _converter;

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var results = new List<bool>();
            var values = _converter.ConvertArgsIncludingOtherTypes(arguments, false);
            var nTrue = 0;
            foreach(var val in values)
            {
                if(val != 0d)
                {
                    nTrue++;
                }
            }
            var result = (System.Math.Abs(nTrue) & 1) != 0;
            return CreateResult(result, DataType.Boolean);
        }
    }
}
