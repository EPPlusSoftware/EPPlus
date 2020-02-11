using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    internal class Concat : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments == null)
            {
                return CreateResult(string.Empty, DataType.String);
            }
            else if(arguments.Count() > 254)
            {
                return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelAddress);
            }
            var sb = new StringBuilder();
            foreach (var arg in arguments)
            {
                if(arg.IsExcelRange)
                {
                    foreach(var cell in arg.ValueAsRangeInfo)
                    {
                        sb.Append(cell.Value);
                    }
                }
                else
                {
                    var v = arg.ValueFirst;
                    if (v != null)
                    {
                        sb.Append(v);
                    }
                }
            }
            var result = sb.ToString();
            if(!string.IsNullOrEmpty(result) && result.Length > 32767)
            {
                return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelAddress);
            }
            return CreateResult(sb.ToString(), DataType.String);
        }
    }
}
