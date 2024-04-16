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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Returns a reference to a cell (or range of cells) for requested rows and columns within a supplied range")]
    internal class Index : ExcelFunction
    {
        public override int ArgumentMinLength => 2;

        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(1, 2);
        }
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arg1 = arguments[0];
            var arg2 = arguments[1];
            if (arg2.DataType == DataType.ExcelError)
            {
                return CompileResult.GetErrorResult(((ExcelErrorValue)arg2.Value).Type);
            }

            int? row = default;

            if (arg2.Value != null)
            {
                row = ArgToInt(arguments, 1, RoundingMethod.Floor);
            }

            int? col = default;
            var colGivenButEmpty = false;
            if(arguments.Count > 2)
            {
                var arg3 = arguments[2];
                if (arg3.DataType == DataType.ExcelError)
                {
                    return CompileResult.GetErrorResult(((ExcelErrorValue)arg3.Value).Type);
                }

                col = ArgToInt(arguments, 2, RoundingMethod.Floor);
                colGivenButEmpty = (arguments[2].Value == null);
            }
            if (!row.HasValue && !col.HasValue) return CreateResult(eErrorType.Value);
            if (arg1.IsExcelRangeOrSingleCell)
            {
                var ri = arg1.ValueAsRangeInfo;
                if(!colGivenButEmpty && !col.HasValue && ri.Size.NumberOfRows > 1 && ri.Size.NumberOfCols > 1)
                {
                    return CreateResult(eErrorType.Ref);
                }
                else if(colGivenButEmpty && row.HasValue)
                {
                    var returnRange = ri.GetOffset(row.Value - 1, 0, row.Value - 1, ri.Size.NumberOfCols - 1);
                    return CreateAddressResult(returnRange, DataType.ExcelRange);
                }
                else if(!row.HasValue && col.HasValue)
                {
                    var returnRange = ri.GetOffset(0, col.Value - 1, ri.Size.NumberOfRows - 1, col.Value - 1);
                    return CreateAddressResult(returnRange, DataType.ExcelRange);
                }
                
                if(ri.Size.NumberOfRows==1 && arguments.Count < 3)
                {
                    var t = row;
                    row = col;
                    col = t;
                }
                if(row==0 || col==0)
                {
                    var range = GetResultRange(row.Value, col.Value, ri);
                    return CreateAddressResult(range, DataType.ExcelRange);
                }
                else
                {
                    return GetResultSingleCell(row != null ? row.Value : 1, col != null ? col.Value : 1, ri);
                }
            }            
            if (arg1.ValueIsExcelError)
            {
                return new CompileResult(arg1.ValueAsExcelErrorValue.Type);
            }
            else
            {
                if(row>1 || col>1)
                {
                    return CompileResult.GetErrorResult(eErrorType.Ref);
                }
                else
                {
                    return CreateResult(arg1.Value, arg1.DataType);
                }
            }
        }
        private static IRangeInfo GetResultRange(int row, int col, IRangeInfo ri)
        {

            return ri.GetOffset(
                row == 0 ? 0 : row-1,
                col == 0 ? 0 : col-1,
                row == 0 ? ri.Size.NumberOfRows - 1 : row - 1,
                col == 0 ? ri.Size.NumberOfCols - 1 : col - 1);                
        }
        private CompileResult GetResultSingleCell(int row, int col, IRangeInfo ri)
        {
            if (row > ri.Address.ToRow - ri.Address.FromRow + 1 ||
                col > ri.Address.ToCol - ri.Address.FromCol + 1)
            {
                return CompileResult.GetErrorResult(eErrorType.Ref);
            }
            var r = row - 1;
            var c = col - 1;

            if (ri.IsInMemoryRange)
            {
                return CompileResultFactory.Create(ri.GetValue(r, c));
            }
            else
            {
                var newRange = ri.GetOffset(r, c, r, c);
                var v = newRange.GetOffset(0, 0);
                return CompileResultFactory.Create(v, newRange.Address);
            }
        }

        public override bool ReturnsReference => true;
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            if (argumentIndex == 0)
            {
                return FunctionParameterInformation.IgnoreAddress;
            }
            return FunctionParameterInformation.IgnoreErrorInPreExecute;
        }));
		/// <summary>
		/// If the function is allowed in a pivot table calculated field
		/// </summary>
		public override bool IsAllowedInCalculatedPivotTableField => false;
	}
}
