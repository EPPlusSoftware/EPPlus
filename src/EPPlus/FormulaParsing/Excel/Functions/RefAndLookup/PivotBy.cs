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
using OfficeOpenXml.FormulaParsing.Ranges;
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
            var result = RenderPivot(rowLevels, colLevels, pivotMap, args, context);

            return CreateDynamicArrayResult(result, DataType.ExcelRange);            
        }

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

                // Bygg radträd och hämta lövet
                var rowLeaf = InsertIntoTree(rowDict, rowOrder, rowKeyParts);

                // Lägg till värden på lövet (precis som BuildGroups gör)
                var leafKey = MakePivotKey(rowKeyParts);
                var row = rowLeaf.Rows.FirstOrDefault(rw => MakePivotKey(rw.KeyParts) == leafKey);
                if (row == null)
                {
                    row = new GroupRow { KeyParts = rowKeyParts };
                    rowLeaf.Rows.Add(row);
                }
                row.Values.Add(vals);

                // Bygg kolumnträd
                InsertIntoTree(colDict, colOrder, colKeyParts);

                // Bygg pivotkartan – sammansatt nyckel per axel för snabb uppslagning
                string rowKey = MakePivotKey(rowKeyParts);

                var colLeaf = InsertIntoTree(colDict, colOrder, colKeyParts);

                // Lägg till värden på kolumnlövet
                var colKey = MakePivotKey(colKeyParts);
                var colRow = colLeaf.Rows.FirstOrDefault(rw => MakePivotKey(rw.KeyParts) == colKey);
                if (colRow == null)
                {
                    colRow = new GroupRow { KeyParts = colKeyParts };
                    colLeaf.Rows.Add(colRow);
                }
                colRow.Values.Add(vals);

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

            // Aggregera subtotaler i radträdet och colträdet.
            AggregateTree(rowLevels, args, context);
            AggregateTree(colLevels, args, context);
        }


        /// <summary>
        /// Infogar en rad med sammansatt nyckel i ett träd.
        /// Extraherad hjälpmetod – identisk logik som i BuildGroups, delad av GROUPBY och PIVOTBY.
        /// </summary>
        protected GroupLevel InsertIntoTree(
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
            return currentLevel; // <-- returnera lövet
        }

        private static string MakePivotKey(object[] parts) =>
            string.Join("\u001F", parts.Select(p => p?.ToString()?.ToLowerInvariant() ?? string.Empty).ToArray());

        private InMemoryRange RenderPivot(
    List<GroupLevel> rowLevels,
    List<GroupLevel> colLevels,
    Dictionary<string, Dictionary<string, List<object[]>>> pivotMap,
    PivotByArgs args,
    ParsingContext context)
        {
            var rowLeaves = CollectLeavesWithPath(rowLevels, new object[0]);
            var colLeaves = CollectLeavesWithPath(colLevels, new object[0]);

            int nRowKeyCols = args.RowFields.Size.NumberOfCols;
            int nColLeaves = colLeaves.Count;
            int nRowLeaves = rowLeaves.Count;
            int nValCols = args.Values.Size.NumberOfCols;

            bool showRowTotal = args.RowTotalDepth != TotalDepthNoTotals;
            bool showColTotal = args.ColTotalDepth != TotalDepthNoTotals;
            bool rowTotalAtTop = args.RowTotalDepth < 0;
            bool colTotalAtLeft = args.ColTotalDepth < 0;

            // Dimensioner
            int dataRows = nRowLeaves;
            int dataCols = nColLeaves;
            int totalRows = 1 + dataRows + (showRowTotal ? 1 : 0);  // rubrik + data + ev. grand total-rad
            int totalCols = nRowKeyCols + dataCols + (showColTotal ? 1 : 0); // radnycklar + data + ev. grand total-kolumn

            var result = new InMemoryRange(totalRows, (short)totalCols);

            // Kolumnindex för grand total-kolumnen
            int grandTotalCol = colTotalAtLeft ? nRowKeyCols : nRowKeyCols + dataCols;
            // Radindex för grand total-raden
            int grandTotalRow = rowTotalAtTop ? 1 : 1 + dataRows;

            // --- Rubrikrad (rad 0) ---
            int colOffset = colTotalAtLeft ? 1 : 0;
            for (int ci = 0; ci < nColLeaves; ci++)
            {
                var colPath = colLeaves[ci].Path;
                result.SetValue(0, nRowKeyCols + colOffset + ci, colPath[colPath.Length - 1]);
            }
            if (showColTotal)
                result.SetValue(0, grandTotalCol, "Total");

            // --- Datarader ---
            int dataRowStart = rowTotalAtTop ? 2 : 1; // hoppa över grand total-raden om den är överst
            for (int ri = 0; ri < nRowLeaves; ri++)
            {
                int outputRow = dataRowStart + ri;
                var rowPath = rowLeaves[ri].Path;
                var rowLeaf = rowLeaves[ri].Leaf;

                // Radnycklar
                for (int k = 0; k < rowPath.Length; k++)
                    result.SetValue(outputRow, k, rowPath[k]);

                string rowKey = MakePivotKey(rowPath);

                // Datavärden per kolumnlöv
                int colIdx = colTotalAtLeft ? nRowKeyCols + 1 : nRowKeyCols;
                for (int ci = 0; ci < nColLeaves; ci++)
                {
                    string colKey = MakePivotKey(colLeaves[ci].Path);

                    object aggregated = null;
                    if (pivotMap.TryGetValue(rowKey, out var colMap) &&
                        colMap.TryGetValue(colKey, out var cellVals))
                    {
                        aggregated = Aggregate(args.Function, cellVals, context,
                            args.Function.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null);
                    }
                    result.SetValue(outputRow, colIdx + ci, aggregated);
                }

                // Grand total-kolumn: SubtotalValues för detta radlöv
                if (showColTotal)
                {
                    var grandTotalVal = rowLeaf.SubtotalValues != null && rowLeaf.SubtotalValues.Count > 0
                        ? rowLeaf.SubtotalValues[0][0]
                        : null;
                    result.SetValue(outputRow, grandTotalCol, grandTotalVal);
                }
            }

            // --- Grand total-rad ---
            if (showRowTotal)
            {
                result.SetValue(grandTotalRow, 0, "Total");
                for (int c = 1; c < nRowKeyCols; c++)
                    result.SetValue(grandTotalRow, c, string.Empty);

                int colIdx = colTotalAtLeft ? nRowKeyCols + 1 : nRowKeyCols;
                for (int ci = 0; ci < nColLeaves; ci++)
                {
                    var colLeaf = colLeaves[ci].Leaf;
                    var grandTotalVal = colLeaf.SubtotalValues != null && colLeaf.SubtotalValues.Count > 0
                        ? colLeaf.SubtotalValues[0][0]
                        : null;
                    result.SetValue(grandTotalRow, colIdx + ci, grandTotalVal);
                }

                // Hörncellen: aggregera hela AllValuesInOrder
                if (showColTotal)
                {
                    var colValues = args.AllValuesInOrder
                        .Select(v => new object[] { v[0] })
                        .ToList();
                    var cornerVal = Aggregate(args.Function, colValues, context,
                        args.Function.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null);
                    result.SetValue(grandTotalRow, grandTotalCol, cornerVal);
                }
            }

            return result;
        }

        private List<LeafWithPath> CollectLeavesWithPath(List<GroupLevel> levels, object[] parentPath)
        {
            var result = new List<LeafWithPath>();
            foreach (var level in levels)
            {
                var path = new object[parentPath.Length + 1];
                for (int i = 0; i < parentPath.Length; i++)
                    path[i] = parentPath[i];
                path[parentPath.Length] = level.Key;

                if (level.IsLeaf)
                    result.Add(new LeafWithPath(level, path));
                else
                    result.AddRange(CollectLeavesWithPath(level.Children, path));
            }
            return result;
        }

        private class LeafWithPath
        {
            public GroupLevel Leaf { get; private set; }
            public object[] Path { get; private set; }

            public LeafWithPath(GroupLevel leaf, object[] path)
            {
                Leaf = leaf;
                Path = path;
            }
        }

    }
}