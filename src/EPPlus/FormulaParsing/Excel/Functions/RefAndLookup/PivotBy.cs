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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.GroupingFunctions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.LookupAndReference,
       EPPlusVersion = "8.6",
       Description = "Allows you to create a summary of your data via a formula. It supports grouping along two axis and aggregating the associated values.")]
    internal class PivotBy : GroupByFunctionBase
    {
        public override int ArgumentMinLength => 3; // kanske ska vara 4
        public override string NamespacePrefix => "_xlfn.";
        public override bool ExecutesLambda => true;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (!TryParsePivotByArgs(arguments, out var args, out var error))
                return error;
            BuildPivotData(args, context, out var rowLevels, out var colLevels, out var pivotMap);
            rowLevels = ApplySort(rowLevels, args);
            colLevels = ApplySort(colLevels, args);

            throw new NotImplementedException();
        }

        // 2. Bygg radgrupperingar  → List<GroupLevel> rowLevels  (återanvänd BuildGroups)
        // 3. Bygg kolumngrupperingar → List<GroupLevel> colLevels  (återanvänd BuildGroups på col_fields)
        // 4. Bygg aggregeringskartan: Dictionary<(rowKey, colKey), object> pivotMap
        // 5. Applicera sortering på båda axlarna (återanvänd ApplySort)
        // 6. Räkna ut dimensioner: nRows, nCols
        // 7. Skriv kolumnhuvuden (ett pass)
        // 8. Skriv datarader rad för rad, slå upp pivotMap[(rowKey, colKey)] per kolumn
        // 9. Skriv grand total-rad/-kolumn

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
                var v = Convert.ToInt32(arguments[3].Value);
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

            // SortOrder for RowFields (optional)
            if (arguments.Count > 6 && arguments[6].Value != null)
                args.RowSortOrders = ParseSortOrderArg(arguments[6]);

            // TotalDepth for columns (optional)
            if (arguments.Count > 7 && arguments[7].Value != null)
            {
                if (!TryParseTotalDepthArg(arguments[7], args.ColFields.Size.NumberOfCols, out int colTotalDepth))
                    return Fail(eErrorType.Value, out error);
                args.ColTotalDepth = colTotalDepth;
            }

            // SortOrder for ColFields (optional)
            if (arguments.Count > 8 && arguments[8].Value != null)
                args.RowSortOrders = ParseSortOrderArg(arguments[8]);

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


        private void BuildPivotData(
            PivotByArgs args,
            ParsingContext context,
            out List<GroupLevel> rowLevels,
            out List<GroupLevel> colLevels,
            out Dictionary<string, Dictionary<string, List<object[]>>> pivotMap)
        {
            var resolvedHeaders = ResolveHeaders(args.Headers, args.Values);
            bool hasHeaders = resolvedHeaders == FieldHeaders.YesAndShow
                           || resolvedHeaders == FieldHeaders.YesAndDontShow;
            int startRow = hasHeaders ? 1 : 0;

            int nRowKeyCols = args.RowFields.Size.NumberOfCols;
            int nColKeyCols = args.ColFields.Size.NumberOfCols;
            int nValCols = args.Values.Size.NumberOfCols;

            var rowDict = new Dictionary<string, GroupLevel>(StringComparer.OrdinalIgnoreCase);
            var rowOrder = new List<string>();
            var colDict = new Dictionary<string, GroupLevel>(StringComparer.OrdinalIgnoreCase);
            var colOrder = new List<string>();
            pivotMap = new Dictionary<string, Dictionary<string, List<object[]>>>(StringComparer.OrdinalIgnoreCase);

            int nRows = args.RowFields.Size.NumberOfRows;

            for (int r = startRow; r < nRows; r++)
            {
                // Filter
                if (args.FilterArray != null)
                {
                    var fv = args.FilterArray.GetOffset(r, 0);
                    if (fv is bool b && !b) continue;
                    if (fv is int i && i == 0) continue;
                }

                // Läs radnycklar och kolumnnycklar
                var rowKeyParts = new object[nRowKeyCols];
                for (int c = 0; c < nRowKeyCols; c++)
                    rowKeyParts[c] = args.RowFields.GetOffset(r, c);

                var colKeyParts = new object[nColKeyCols];
                for (int c = 0; c < nColKeyCols; c++)
                    colKeyParts[c] = args.ColFields.GetOffset(r, c);

                var vals = new object[nValCols];
                for (int c = 0; c < nValCols; c++)
                    vals[c] = args.Values.GetOffset(r, c);

                // Bygg radträd
                InsertIntoTree(rowDict, rowOrder, rowKeyParts);

                // Bygg kolumnträd
                InsertIntoTree(colDict, colOrder, colKeyParts);

                // Bygg pivotkartan – sammansatt nyckel per axel för snabb uppslagning
                string rowKey = MakePivotKey(rowKeyParts);
                string colKey = MakePivotKey(colKeyParts);

                if (!pivotMap.TryGetValue(rowKey, out var colMap))
                {
                    colMap = new Dictionary<string, List<object[]>>(StringComparer.OrdinalIgnoreCase);
                    pivotMap[rowKey] = colMap;
                }
                if (!colMap.TryGetValue(colKey, out var cellVals))
                {
                    cellVals = new List<object[]>();
                    colMap[colKey] = cellVals;
                }
                cellVals.Add(vals);
                args.AllValuesInOrder.Add(vals);
            }

            rowLevels = BuildOrderedTree(rowDict, rowOrder);
            colLevels = BuildOrderedTree(colDict, colOrder);

            // Aggregera subtotaler i radträdet (återanvänder AggregateTree från basklassen)
            AggregateTree(rowLevels, args, context);
        }


        /// <summary>
        /// Infogar en rad med sammansatt nyckel i ett träd.
        /// Extraherad hjälpmetod – identisk logik som i BuildGroups, delad av GROUPBY och PIVOTBY.
        /// </summary>
        protected void InsertIntoTree(
            Dictionary<string, GroupLevel> rootDict,
            List<string> rootOrder,
            object[] keyParts)
        {
            var currentDict = rootDict;
            var currentOrder = rootOrder;
            GroupLevel currentLevel = null;

            for (int depth = 0; depth < keyParts.Length; depth++)
            {
                var keyStr = (keyParts[depth]?.ToString() ?? string.Empty).ToLowerInvariant();
                if (!currentDict.TryGetValue(keyStr, out currentLevel))
                {
                    currentLevel = new GroupLevel { Key = keyParts[depth] };
                    currentDict[keyStr] = currentLevel;
                    currentOrder.Add(keyStr);
                }

                if (depth < keyParts.Length - 1)
                {
                    if (currentLevel.ChildDict == null)
                    {
                        currentLevel.ChildDict = new Dictionary<string, GroupLevel>(StringComparer.OrdinalIgnoreCase);
                        currentLevel.ChildOrder = new List<string>();
                    }
                    currentDict = currentLevel.ChildDict;
                    currentOrder = currentLevel.ChildOrder;
                }
            }
        }

        private static string MakePivotKey(object[] parts) =>
            string.Join("\u001F", parts.Select(p => p?.ToString()?.ToLowerInvariant() ?? string.Empty).ToArray());
    }
}
