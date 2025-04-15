/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/3/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    internal abstract class ChooseFunction : ExcelFunction
    {
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 2;
        protected List<int> GetChooseColumns(IList<FunctionArgument> arguments, out eErrorType? ev)
        {
            var cols = new List<int>();
            for (var x = 1; x < arguments.Count; x++)
            {
                if (arguments[x].IsExcelRange)
                {
                    var range = arguments[x].ValueAsRangeInfo;
                    if (range.Size.NumberOfRows > 1 && range.Size.NumberOfCols > 1)
                    {
                        ev = eErrorType.Value;
                    }
                    for (int r = 0; r < range.Size.NumberOfRows; r++)
                    {
                        for (int c = 0; c < range.Size.NumberOfCols; c++)
                        {
                            var v = range.GetOffset(r, c);
                            if (v is ExcelErrorValue error)
                            {
                                ev = error.Type;
                                return null;
                            }
                            var d = ConvertUtil.GetValueDouble(v, false, true);
                            if (double.IsNaN(d))
                            {
                                ev = eErrorType.Value;
                            }
                            else
                            {

                                cols.Add((int)Math.Truncate(d));
                            }
                        }
                    }
                }
                else
                {
                    var c = ArgToInt(arguments, x, out ExcelErrorValue e1);
                    if (e1 != null)
                    {
                        ev = e1.Type;
                        return null;
                    }
                    cols.Add(c);
                }
            }
            ev = null;
            return cols;
        }

        /// <summary>
        /// If the function is allowed in a pivot table calculated field
        /// </summary>
        public override bool IsAllowedInCalculatedPivotTableField => false;
    }
}
