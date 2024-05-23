using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal abstract class ToRowColBase : ExcelFunction
    {
        protected List<object> GetItemsFromRange(IRangeInfo range, int ignore, bool scanByColumn)
        {
            var result = new List<object>();
            var maxX = scanByColumn ? range.Size.NumberOfCols : range.Size.NumberOfRows;
            var maxy = scanByColumn ? range.Size.NumberOfRows : range.Size.NumberOfCols;
            for (var x = 0; x < maxX; x++)
            {
                for (short y = 0; y < maxy; y++)
                {
                    var v = scanByColumn ? range.GetOffset(y, x) : range.GetOffset(x, y);
                    if ((ignore == 1 || ignore == 3) && v == null)
                    {
                        continue;
                    }
                    else if ((ignore == 2 || ignore == 3) && ExcelErrorValue.IsErrorValue(v?.ToString()))
                    {
                        continue;
                    }
                    result.Add(v);
                }
            }
            return result;
        }
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
