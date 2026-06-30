using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    internal abstract class RegexFunctionBase : ExcelFunction
    {
        protected static string GetValue(
                        IRangeInfo range,
                        FunctionArgument scalar,
                        int nRows, int nCols,
                        int row, int col)
        {
            if (range == null)                
                return scalar.Value?.ToString();
            
            int r = nRows == 1 ? 0 : row;
            int c = nCols == 1 ? 0 : col;

            if (r >= nRows || c >= nCols)
                return null;

            return range.GetOffset(r, c)?.ToString();
        }

        protected static short ExpandedSize(int a, int b)
        {
            if (a == 1) return (short)b;
            if (b == 1) return (short)a;
            return (short)Math.Max(a, b);  
        }
    }
}
