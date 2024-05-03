using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    internal abstract class ExcelFunctionTextBase : ExcelFunction
    {
        protected string ArgDelimiterCollectionToString(IList<FunctionArgument> arguments, int index, out CompileResult error)
        {
            var obj = arguments[index].ValueToList;
            var str = string.Empty;
            foreach (var b in obj)
            {
                if(string.IsNullOrEmpty(b.ToString()))
                {
                    error = CompileResult.GetErrorResult(eErrorType.Value);
                    return null;
                }
                str += b.ToString();
            }
            error = null;
            return str;
        }
    }
}
