/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  13/4/2026         EPPlus Software AB           EPPlus v8.6
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.GroupingFunctions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.PivotBy
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.LookupAndReference,
       EPPlusVersion = "8.6",
       Description = "Allows you to create a summary of your data via a formula. It supports grouping along two axis and aggregating the associated values.")]
    internal partial class PivotBy : GroupByFunctionBase
    {
        public override int ArgumentMinLength => 3; 
        public override string NamespacePrefix => "_xlfn.";
        public override bool ExecutesLambda => true;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (!TryParsePivotByArgs(arguments, out var args, out var error))
                return error;
            BuildPivotData(args, context, out var rowLeaves, out var colLeaves, out var pivotMap);
            rowLeaves = ApplyRowSort(rowLeaves, args);
            colLeaves = ApplyColSort(colLeaves, args);
            var result = RenderPivot(rowLeaves, colLeaves, pivotMap, args, context);

            return CreateDynamicArrayResult(result, DataType.ExcelRange);            
        }
            
        protected bool TryParsePivotByArgs(IList<FunctionArgument> arguments,
            out PivotByArgs args,
            out CompileResult error)
        {
            args = new PivotByArgs();
            error = null;

            args.RowFields = arguments[0].ValueAsRangeInfo;
            args.ColFields = arguments[1].ValueAsRangeInfo;
            args.Values = arguments[2].ValueAsRangeInfo;

            if (args.RowFields.Size.NumberOfRows != args.Values.Size.NumberOfRows)
                return Fail(eErrorType.Value, out error);

            if (!TryParseFunctionArg(arguments[3], args.Functions, out LambdaCalculator function, out FunctionLayout layout))
                return Fail(eErrorType.Value, out error);

            args.Function = function;
            args.FunctionLayout = layout;

            if (arguments.Count > 4 && arguments[4].Value != null)
            {
                var v = Convert.ToInt32(arguments[4].Value);
                if (!Enum.IsDefined(typeof(FieldHeaders), v))
                    return Fail(eErrorType.Value, out error);
                args.Headers = (FieldHeaders)v;
            }

            // Total depth for rows (optional)
            if (arguments.Count > 5 && arguments[5].Value != null)
            {
                if (!TryParseTotalDepthArg(arguments[5], args.RowFields.Size.NumberOfCols, out int rowTotalDepth))
                    return Fail(eErrorType.Value, out error);
                args.RowTotalDepth = rowTotalDepth;
            }
            else
            {
                args.RowTotalDepth = TotalDepthGrandOnly; // default = 1
            }

            // SortOrder for RowFields (optional)
            if (arguments.Count > 6 && arguments[6].Value != null)
                args.RowSortOrders = ParseSortOrderArg(arguments[6]);

            // TotalDepth for columns (optional, default = 1)
            if (arguments.Count > 7 && arguments[7].Value != null)
            {
                if (!TryParseTotalDepthArg(arguments[7], args.ColFields.Size.NumberOfCols, out int colTotalDepth))
                    return Fail(eErrorType.Value, out error);
                args.ColTotalDepth = colTotalDepth;
            }
            else
            {
                args.ColTotalDepth = TotalDepthGrandOnly; // default = 1
            }

            // SortOrder for ColFields (optional)
            if (arguments.Count > 8 && arguments[8].Value != null)
                args.ColSortOrders = ParseSortOrderArg(arguments[8]);

            if (arguments.Count > 9 && arguments[9].IsExcelRange)
                args.FilterArray = arguments[9].ValueAsRangeInfo;

            // RelativeTo (optional)
            if (arguments.Count > 10 && arguments[10].Value != null)
            {
                var v = Convert.ToInt32(arguments[10].Value);
                if (!Enum.IsDefined(typeof(RelativeTo), v))
                    return Fail(eErrorType.Value, out error);
                args.RelativeTo = (RelativeTo)v;
            }

            return true;
        }
    }
}