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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.GroupingFunctions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.PivotBy
{
    internal partial class PivotBy
    {
        private Dictionary<string, List<object[]>> _colTotalsCache;
        private Dictionary<string, List<object[]>> _rowTotalsCache;

        private List<object[]> ResolveRelativeToValues(
            RelativeTo relativeTo,
            LeafWithPath colLeaf,
            List<LeafWithPath> colLeaves,
            string rowKey,
            object[] rowPath,
            List<LeafWithPath> rowLeaves,
            Dictionary<string, Dictionary<string, List<object[]>>> pivotMap,
            PivotByArgs args)
        {
            switch (relativeTo)
            {
case RelativeTo.ColumnTotals: // 0 — default
    {
        string colKey = colLeaf.PivotKey;

        if (_colTotalsCache != null && _colTotalsCache.TryGetValue(colKey, out var cached))
            return cached;

        var list = pivotMap.Values
            .SelectMany(cm => cm.TryGetValue(colKey, out var cv)
                ? cv
                : Enumerable.Empty<object[]>())
            .ToList();

        if (_colTotalsCache != null)
            _colTotalsCache[colKey] = list;
        return list;
    }
case RelativeTo.RowTotals: // 1
    {
        if (_rowTotalsCache != null && _rowTotalsCache.TryGetValue(rowKey, out var cached))
            return cached;

        if (!pivotMap.TryGetValue(rowKey, out var colMap))
            return null;

        var list = colMap.Values
            .SelectMany(vals => vals)
            .ToList();

        if (_rowTotalsCache != null)
            _rowTotalsCache[rowKey] = list;
        return list;
    }
                case RelativeTo.GrandTotals: // 2
                    {
                        return args.AllValuesInOrder;
                    }
                case RelativeTo.ParentColTotal:
                    {
                        if (!pivotMap.TryGetValue(rowKey, out var colMap))
                            return null;

                        if (colLeaf.Path.Length <= 1)
                        {
                            return colMap.Values
                                .SelectMany(vals => vals)
                                .ToList();
                        }
                       
                        var parentPrefix = GetParentPrefix(colLeaf.Path);
                        var siblingLeaves = colLeaves
                            .Where(l => HasParentPrefix(l.Path, parentPrefix))
                            .ToList();

                        return siblingLeaves
                            .Where(leaf => colMap.ContainsKey(leaf.PivotKey))
                            .SelectMany(leaf => colMap[leaf.PivotKey])
                            .ToList();
                    }
                case RelativeTo.ParentRowTotal: // 4
                    {
                        string colKey = colLeaf.PivotKey;
                    
                        if (rowPath == null || rowPath.Length <= 1)
                        {
                            return pivotMap.Values
                                .SelectMany(cm => cm.TryGetValue(colKey, out var cv)
                                    ? cv
                                    : Enumerable.Empty<object[]>())
                                .ToList();
                        }

                        var parentPrefix = GetParentPrefix(rowPath);
                        return rowLeaves
                            .Where(rl => HasParentPrefix(rl.Path, parentPrefix))
                            .Select(rl => rl.PivotKey)
                            .Where(rk => pivotMap.ContainsKey(rk))
                            .SelectMany(rk => pivotMap[rk].TryGetValue(colKey, out var cv)
                                ? cv
                                : Enumerable.Empty<object[]>())
                            .ToList();
                    }
                default:
                    return null;
            }
        }

        private static string[] GetParentPrefix(object[] path)
        {
            int parentDepth = path.Length - 1;
            var prefix = new string[parentDepth];
            for (int i = 0; i < parentDepth; i++)
                prefix[i] = path[i]?.ToString()?.ToLowerInvariant() ?? string.Empty;
            return prefix;
        }

        private static bool HasParentPrefix(object[] path, string[] parentPrefix)
        {
            if (path.Length < parentPrefix.Length) return false;
            for (int i = 0; i < parentPrefix.Length; i++)
            {
                var pv = path[i]?.ToString()?.ToLowerInvariant() ?? string.Empty;
                if (pv != parentPrefix[i]) return false;
            }
            return true;
        }

        /// <summary>
        /// Returns the row parent group's values across ALL columns.
        /// Used by PERCENTOF when RelativeTo=ParentRowTotal in the row-total cell:
        /// the denominator is the parent group's whole-row total summed over every column.
        /// Returns null when the row has no real parent (single row level), letting the caller
        /// fall back to a sensible default.
        /// </summary>
        private static List<object[]> GetParentRowGroupAllValues(
            object[] rowPath,
            List<LeafWithPath> rowLeaves)
        {
            if (rowPath == null || rowPath.Length <= 1) return null;

            var parentPrefix = GetParentPrefix(rowPath);
            return rowLeaves
                .Where(rl => HasParentPrefix(rl.Path, parentPrefix))
                .SelectMany(rl => rl.Leaf.Rows.SelectMany(row => row.Values))
                .ToList();
        }

        /// <summary>
        /// Resolves the PERCENTOF denominator for a row-subtotal cell when RelativeTo=ParentColTotal.
        /// The denominator must be the row group's values restricted to the column's parent group -
        /// NOT all values for the row group across every column.
        ///
        /// For a column subtotal entry (which IS a parent group already), we use any of its
        /// GroupLeaves to derive the parent prefix. For a regular leaf entry we use its own path.
        /// When there is only one column level, no real parent exists and we fall back to
        /// summing every column for the row group.
        /// </summary>
        private static List<object[]> ResolveSubtotalParentColGroupValues(
            ColEntry entry,
            List<LeafWithPath> colLeaves,
            List<string> groupRowKeys,
            Dictionary<string, Dictionary<string, List<object[]>>> pivotMap)
        {
            var colPath = entry.IsSubtotal
                ? entry.GroupLeaves[0].Path
                : entry.Leaf.Path;

            if (colPath.Length <= 1)
            {
                return groupRowKeys
                    .SelectMany(rk => pivotMap.TryGetValue(rk, out var cm)
                        ? cm.Values.SelectMany(v => v)
                        : Enumerable.Empty<object[]>())
                    .ToList();
            }

            var parentPrefix = GetParentPrefix(colPath);
            var siblingColLeaves = colLeaves
                .Where(cl => HasParentPrefix(cl.Path, parentPrefix))
                .ToList();

            return groupRowKeys
                .SelectMany(rk => pivotMap.TryGetValue(rk, out var cm)
                    ? siblingColLeaves
                        .Where(cl => cm.ContainsKey(cl.PivotKey))
                        .SelectMany(cl => cm[cl.PivotKey])
                    : Enumerable.Empty<object[]>())
                .ToList();
        }

        /// <summary>
        /// Returns the denominator values list for PERCENTOF in the row-total cell of a data row.
        /// - RowTotals / ParentColTotal: the row's own values (cell / row total = 1)
        /// - ParentRowTotal: the row parent group's values across all columns
        /// - ColumnTotals / GrandTotals (default): the whole dataset
        /// </summary>
        private static List<object[]> ResolveRowTotalCellRelativeToVals(
            PivotByArgs args,
            List<object[]> rowAllVals,
            object[] rowPath,
            List<LeafWithPath> rowLeaves)
        {
            switch (args.RelativeTo)
            {
                case RelativeTo.RowTotals:
                case RelativeTo.ParentColTotal:
                    return rowAllVals;
                case RelativeTo.ParentRowTotal:
                    return GetParentRowGroupAllValues(rowPath, rowLeaves) ?? args.AllValuesInOrder;
                default: // ColumnTotals, GrandTotals
                    return args.AllValuesInOrder;
            }
        }

        private List<object[]> ResolveRelativeToValuesForTotal(
            RelativeTo relativeTo,
            LeafWithPath colLeaf,
            List<LeafWithPath> colLeaves,
            Dictionary<string, Dictionary<string, List<object[]>>> pivotMap,
            PivotByArgs args)
        {
            switch (relativeTo)
            {
                case RelativeTo.RowTotals:
                    {
                        return pivotMap.Values
                            .SelectMany(colMap => colMap.Values.SelectMany(v => v))
                            .ToList();
                    }
                case RelativeTo.GrandTotals: 
                    {
                        return args.AllValuesInOrder;
                    }
                case RelativeTo.ParentColTotal:
                    {
                        if (colLeaf.Path.Length <= 1)
                        {
                            return pivotMap.Values
                                .SelectMany(colMap => colMap.Values.SelectMany(v => v))
                                .ToList();
                        }

                        var parentPrefix = GetParentPrefix(colLeaf.Path);
                        var siblingLeaves = colLeaves
                            .Where(l => HasParentPrefix(l.Path, parentPrefix))
                            .ToList();

                        return pivotMap.Values
                            .SelectMany(cm => siblingLeaves
                                .Where(leaf => cm.ContainsKey(MakePivotKey(leaf.Path)))
                                .SelectMany(leaf => cm[leaf.PivotKey]))
                            .ToList();
                    }
                case RelativeTo.ParentRowTotal: 
                    {
                        string colKey = colLeaf.PivotKey;
                        var colVals = pivotMap.Values
                            .SelectMany(cm => cm.TryGetValue(colKey, out var cv)
                                ? cv
                                : Enumerable.Empty<object[]>())
                            .ToList();
                        return colVals;
                    }
                default:
                    return null;
            }
        }

        private static string MakePivotKey(object[] parts) =>
            string.Join("\u001F", parts.Select(p => p?.ToString()?.ToLowerInvariant() ?? string.Empty).ToArray());



        private InMemoryRange RenderPivot(
            List<LeafWithPath> rowLeaves,
            List<LeafWithPath> colLeaves,
            Dictionary<string, Dictionary<string, List<object[]>>> pivotMap,
            PivotByArgs args,
            ParsingContext context)
        {
            if (args.Functions.Any(f => f.EtaFunction?.Name == "PERCENTOF"))
            {
                _colTotalsCache = new Dictionary<string, List<object[]>>(StringComparer.OrdinalIgnoreCase);
                _rowTotalsCache = new Dictionary<string, List<object[]>>(StringComparer.OrdinalIgnoreCase);
            }
            else
            {
                _colTotalsCache = null;
                _rowTotalsCache = null;
            }

            int nRowKeyCols = args.RowFields.Size.NumberOfCols;
            int nColKeyRows = args.ColFields.Size.NumberOfCols;
            int nRowLeaves = rowLeaves.Count;
            int nFunctions = args.Functions.Count;
            bool isVStack = args.FunctionLayout == FunctionLayout.Vertical;
            bool isHStack = args.FunctionLayout == FunctionLayout.Horizontal;

            bool showRowTotal = args.RowTotalDepth != TotalDepthNoTotals;
            bool showColTotal = args.ColTotalDepth != TotalDepthNoTotals;
            bool rowTotalAtTop = args.RowTotalDepth < 0;
            bool colTotalAtLeft = args.ColTotalDepth < 0;
            int colSubtotalDepth = Math.Abs(args.ColTotalDepth);
            bool showColSubtotals = colSubtotalDepth > 1;
            int rowSubtotalDepth = Math.Abs(args.RowTotalDepth);
            bool showRowSubtotals = rowSubtotalDepth > 1;

            var resolvedHeaders = ResolveHeaders(args.Headers, args.Values);
            bool showFieldHeaders = resolvedHeaders == FieldHeaders.YesAndShow;
            int fieldHeaderRows = showFieldHeaders ? 1 : 0;
            int headerDataRows = showFieldHeaders ? 1 : 0;
            int functionHeaderRows = isHStack ? 1 : 0;
            int functionNameCol = isVStack ? nRowKeyCols : -1;
            int functionColOffset = isVStack ? 1 : 0;

            var colGroups = colLeaves
                .GroupBy(l => l.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty)
                .ToList();

            var colEntries = new List<ColEntry>();
            foreach (var group in colGroups)
            {
                var groupLeaves = group.ToList();
                if (showColSubtotals && colTotalAtLeft)
                    colEntries.Add(new ColEntry { IsSubtotal = true, GroupKey = group.Key, GroupLeaves = groupLeaves });
                foreach (var leaf in groupLeaves)
                    colEntries.Add(new ColEntry { IsSubtotal = false, Leaf = leaf });
                if (showColSubtotals && !colTotalAtLeft)
                    colEntries.Add(new ColEntry { IsSubtotal = true, GroupKey = group.Key, GroupLeaves = groupLeaves });
            }

            var rowGroups = rowLeaves
                .GroupBy<LeafWithPath, string>(l => string.Join("\u001F",
                    l.Path.Take(rowSubtotalDepth - 1)
                          .Select(p => p?.ToString()?.ToLowerInvariant() ?? string.Empty)
                          .ToArray()))
                .ToList();

            int colsPerEntry = isHStack ? nFunctions : 1;
            int nDataCols = colEntries.Count * colsPerEntry;
            int nTotalCols = showColTotal ? colsPerEntry : 0;
            int totalCols = nRowKeyCols + functionColOffset + nDataCols + nTotalCols;

            int rowsPerLeaf = isVStack ? nFunctions : 1;
            int rowsPerTotal = isVStack ? nFunctions : 1;
            int subtotalRowCount = showRowSubtotals ? rowGroups.Count * rowsPerLeaf : 0;

            int totalRows = fieldHeaderRows + nColKeyRows + functionHeaderRows + headerDataRows
                          + nRowLeaves * rowsPerLeaf
                          + subtotalRowCount
                          + (showRowTotal ? rowsPerTotal : 0);

            int dataRowStart = fieldHeaderRows + nColKeyRows + functionHeaderRows + headerDataRows
                             + (rowTotalAtTop ? rowsPerTotal : 0);
            int grandTotalRow = fieldHeaderRows + nColKeyRows + functionHeaderRows + headerDataRows
                              + (rowTotalAtTop ? 0 : nRowLeaves * rowsPerLeaf + subtotalRowCount);

            var result = new InMemoryRange(totalRows, (short)totalCols);

            int dataColStart = nRowKeyCols + functionColOffset;
            int grandTotalCol = colTotalAtLeft ? dataColStart : dataColStart + nDataCols;
            int colOffset = colTotalAtLeft ? colsPerEntry : 0;

            // --- Headerrows ---            
            if (showFieldHeaders)
            {
                for (int i = 0; i < totalCols; i++)
                    result.SetValue(0, i, string.Empty);

                for (int i = 0; i < args.ColFields.Size.NumberOfCols; i++)
                    result.SetValue(0, dataColStart + colOffset + i, args.ColFields.GetOffset(0, i));
            }
            
            for (int level = 0; level < nColKeyRows; level++)
            {
                int outputLevel = fieldHeaderRows + level;

                for (int i = 0; i < nRowKeyCols + functionColOffset; i++)
                    result.SetValue(outputLevel, i, string.Empty);

                int col = dataColStart + colOffset;
                foreach (var entry in colEntries)
                {
                    var val = entry.IsSubtotal
                        ? (level == 0 ? entry.GroupLeaves[0].Path[0] : (object)string.Empty)
                        : (level < entry.Leaf.Path.Length ? entry.Leaf.Path[level] : (object)string.Empty);

                    for (int f = 0; f < colsPerEntry; f++)
                    {
                        result.SetValue(outputLevel, col, val);
                        col++;
                    }
                }

                if (showColTotal)
                {
                    string colTotalLabel = Math.Abs(args.ColTotalDepth) > 1 ? "Grand Total" : "Total";
                    int totalCol = grandTotalCol;
                    for (int f = 0; f < colsPerEntry; f++)
                    {
                        result.SetValue(outputLevel, totalCol, level == 0 ? colTotalLabel : string.Empty);
                        totalCol++;
                    }
                }
            }

            // --- HSTACK ---
            if (isHStack)
            {
                var functionNames = ResolveFunctionHeaders(args.Functions);
                int functionHeaderRow = fieldHeaderRows + nColKeyRows;

                for (int i = 0; i < nRowKeyCols; i++)
                    result.SetValue(functionHeaderRow, i, string.Empty);

                int col = dataColStart + colOffset;
                foreach (var entry in colEntries)
                {
                    foreach (var name in functionNames)
                    {
                        result.SetValue(functionHeaderRow, col, name);
                        col++;
                    }
                }
                if (showColTotal)
                {
                    int totalCol = grandTotalCol;
                    foreach (var name in functionNames)
                    {
                        result.SetValue(functionHeaderRow, totalCol, name);
                        totalCol++;
                    }
                }
            }

            // --- Header-datarad ---
            if (showFieldHeaders)
            {
                int headerDataRow = fieldHeaderRows + nColKeyRows + functionHeaderRows;

                for (int i = 0; i < nRowKeyCols; i++)
                    result.SetValue(headerDataRow, i, args.RowFields.GetOffset(0, i));
                if (functionColOffset > 0)
                    result.SetValue(headerDataRow, nRowKeyCols, string.Empty);

                var headerValue = args.Values.GetOffset(0, 0);
                int col = dataColStart + colOffset;
                foreach (var entry in colEntries)
                {
                    for (int fc = 0; fc < colsPerEntry; fc++)
                    {
                        result.SetValue(headerDataRow, col, entry.IsSubtotal ? string.Empty : headerValue);
                        col++;
                    }
                }

                if (showColTotal)
                {
                    int totalCol = grandTotalCol;
                    for (int fc = 0; fc < colsPerEntry; fc++)
                    {
                        result.SetValue(headerDataRow, totalCol, headerValue);
                        totalCol++;
                    }
                }
            }

            if (rowTotalAtTop && showRowTotal)
                WriteGrandTotalRow(result, fieldHeaderRows + nColKeyRows + functionHeaderRows + headerDataRows,
                                   colEntries, colLeaves, pivotMap, args, context,
                                   nRowKeyCols, functionNameCol, dataColStart, colOffset,
                                   grandTotalCol, showColTotal, colsPerEntry, isVStack, functionColOffset);

            // --- Datarows ---
            var functionNames2 = ResolveFunctionHeaders(args.Functions);
            int currentOutputRow = dataRowStart;

            foreach (var rowGroup in rowGroups)
            {
                var groupLeaves = rowGroup.ToList();
                var groupKeyParts = showRowSubtotals
                    ? groupLeaves[0].Path.Take(rowSubtotalDepth - 1).ToArray()
                    : null;

                if (showRowSubtotals && rowTotalAtTop)
                {
                    for (int fi = 0; fi < rowsPerLeaf; fi++)
                    {
                        WriteRowSubtotalRow(
                            result, currentOutputRow + fi,
                            groupKeyParts, groupLeaves,
                            colEntries, colLeaves,
                            pivotMap, args, context,
                            nRowKeyCols, functionNameCol,
                            dataColStart, colOffset,
                            grandTotalCol,
                            showColTotal, isVStack,
                            args.Functions[fi], functionNames2[fi]);
                    }
                    currentOutputRow += rowsPerLeaf;
                }

                foreach (var rowLeafEntry in groupLeaves)
                {
                    var rowPath = rowLeafEntry.Path;
                    var rowLeaf = rowLeafEntry.Leaf;
                    string rowKey = rowLeafEntry.PivotKey;

                    for (int fi = 0; fi < nFunctions; fi++)
                    {
                        int outputRow = currentOutputRow + fi;
                        var f = args.Functions[fi];

                        for (int k = 0; k < nRowKeyCols; k++)
                            result.SetValue(outputRow, k, k < rowPath.Length ? rowPath[k] : string.Empty);

                        if (isVStack)
                            result.SetValue(outputRow, functionNameCol, functionNames2[fi]);

                        int col = dataColStart + colOffset;
                        foreach (var entry in colEntries)
                        {
                            if (entry.IsSubtotal)
                            {
                                var groupVals = entry.GroupLeaves
                                    .SelectMany(l =>
                                    {
                                        var ck = l.PivotKey;
                                        if (pivotMap.TryGetValue(rowKey, out var cm) && cm.TryGetValue(ck, out var cv))
                                            return cv;
                                        return Enumerable.Empty<object[]>();
                                    }).ToList();

                                if (isHStack)
                                {
                                    foreach (var func in args.Functions)
                                    {
                                        object val = groupVals.Count > 0
                                            ? Aggregate(func, groupVals, context,
                                                func.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null)
                                            : null;
                                        result.SetValue(outputRow, col, val ?? string.Empty);
                                        col++;
                                    }
                                }
                                else
                                {
                                    object val = groupVals.Count > 0
                                        ? Aggregate(f, groupVals, context,
                                            f.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null)
                                        : null;
                                    result.SetValue(outputRow, col, val ?? string.Empty);
                                    col++;
                                }
                            }
                            else
                            {
                                string colKey = entry.Leaf.PivotKey;

                                if (isHStack)
                                {
                                    foreach (var func in args.Functions)
                                    {
                                        object aggregated = null;
                                        if (pivotMap.TryGetValue(rowKey, out var colMap) &&
                                            colMap.TryGetValue(colKey, out var cellVals))
                                        {
                                            var relativeToVals = func.EtaFunction?.Name == "PERCENTOF"
                                                ? ResolveRelativeToValues(args.RelativeTo, entry.Leaf, colLeaves, rowKey, rowPath, rowLeaves, pivotMap, args) ?? args.AllValuesInOrder
                                                : args.AllValuesInOrder;
                                            aggregated = Aggregate(func, cellVals, context, relativeToVals);
                                        }
                                        result.SetValue(outputRow, col, aggregated ?? string.Empty);
                                        col++;
                                    }
                                }
                                else
                                {
                                    object aggregated = null;
                                    if (pivotMap.TryGetValue(rowKey, out var colMap) &&
                                        colMap.TryGetValue(colKey, out var cellVals))
                                    {
                                        var relativeToVals = f.EtaFunction?.Name == "PERCENTOF"
                                            ? ResolveRelativeToValues(args.RelativeTo, entry.Leaf, colLeaves, rowKey, rowPath, rowLeaves, pivotMap, args) ?? args.AllValuesInOrder
                                            : args.AllValuesInOrder;
                                        aggregated = Aggregate(f, cellVals, context, relativeToVals);
                                    }
                                    result.SetValue(outputRow, col, aggregated ?? string.Empty);
                                    col++;
                                }
                            }
                        }

                        if (showColTotal)
                        {
                            var rowAllVals = rowLeaf.Rows.SelectMany(r => r.Values).ToList();
                            int totalCol = grandTotalCol;
                            if (isHStack)
                            {
                                foreach (var func in args.Functions)
                                {
                                    var relVals = args.AllValuesInOrder;
                                    if (func.EtaFunction?.Name == "PERCENTOF")
                                        relVals = ResolveRowTotalCellRelativeToVals(args, rowAllVals, rowPath, rowLeaves);
                                    var totalVal = Aggregate(func, rowAllVals, context, relVals);
                                    result.SetValue(outputRow, totalCol, totalVal ?? string.Empty);
                                    totalCol++;
                                }
                            }
                            else
                            {
                                var relVals = args.AllValuesInOrder;
                                if (f.EtaFunction?.Name == "PERCENTOF")
                                    relVals = ResolveRowTotalCellRelativeToVals(args, rowAllVals, rowPath, rowLeaves);
                                var totalVal = Aggregate(f, rowAllVals, context, relVals);
                                result.SetValue(outputRow, totalCol, totalVal ?? string.Empty);
                            }
                        }

                        if (isHStack) break;
                    }

                    currentOutputRow += rowsPerLeaf;
                }

                if (showRowSubtotals && !rowTotalAtTop)
                {
                    for (int fi = 0; fi < rowsPerLeaf; fi++)
                    {
                        WriteRowSubtotalRow(
                            result, currentOutputRow + fi,
                            groupKeyParts, groupLeaves,
                            colEntries, colLeaves,
                            pivotMap, args, context,
                            nRowKeyCols, functionNameCol,
                            dataColStart, colOffset,
                            grandTotalCol,
                            showColTotal, isVStack,
                            args.Functions[fi], functionNames2[fi]);
                    }
                    currentOutputRow += rowsPerLeaf;
                }
            }
            
            if (!rowTotalAtTop && showRowTotal)
                WriteGrandTotalRow(result, grandTotalRow, colEntries, colLeaves, pivotMap, args, context,
                                   nRowKeyCols, functionNameCol, dataColStart, colOffset,
                                   grandTotalCol, showColTotal, colsPerEntry, isVStack, functionColOffset);

            return result;
        }


        private void WriteRowSubtotalRow(
             InMemoryRange result,
             int outputRow,
             object[] groupKeyParts,
             List<LeafWithPath> groupLeaves,
             List<ColEntry> colEntries,
             List<LeafWithPath> colLeaves,
             Dictionary<string, Dictionary<string, List<object[]>>> pivotMap,
             PivotByArgs args,
             ParsingContext context,
             int nRowKeyCols,
             int functionNameCol,
             int dataColStart,
             int colOffset,
             int grandTotalCol,
             bool showColTotal,
             bool isVStack,
             LambdaCalculator f,
             string functionName)
        {
            for (int k = 0; k < nRowKeyCols; k++)
                result.SetValue(outputRow, k, k < groupKeyParts.Length ? groupKeyParts[k] : string.Empty);

            if (isVStack)
                result.SetValue(outputRow, functionNameCol, functionName);

            var groupRowKeys = groupLeaves.Select(l => l.PivotKey).ToList();

            int col = dataColStart + colOffset;
            foreach (var entry in colEntries)
            {
                List<object[]> cellVals;
                if (entry.IsSubtotal)
                {
                    cellVals = groupRowKeys
                        .SelectMany(rk =>
                        {
                            if (!pivotMap.TryGetValue(rk, out var cm)) return Enumerable.Empty<object[]>();
                            return entry.GroupLeaves
                                .SelectMany(gl => cm.TryGetValue(gl.PivotKey, out var cv)
                                    ? cv : Enumerable.Empty<object[]>());
                        }).ToList();
                }
                else
                {
                    string colKey = entry.Leaf.PivotKey;
                    cellVals = groupRowKeys
                        .SelectMany(rk =>
                        {
                            if (!pivotMap.TryGetValue(rk, out var cm)) return Enumerable.Empty<object[]>();
                            return cm.TryGetValue(colKey, out var cv) ? cv : Enumerable.Empty<object[]>();
                        }).ToList();
                }

                if (cellVals.Count > 0)
                {
                    List<object[]> relativeToVals = null;
                    if (f.EtaFunction?.Name == "PERCENTOF")
                    {
                        var effectiveRelativeTo = args.RelativeTo == RelativeTo.ColumnTotals
                            ? RelativeTo.ParentRowTotal
                            : args.RelativeTo;

                        relativeToVals = effectiveRelativeTo switch
                        {
                            RelativeTo.RowTotals =>
                                groupRowKeys
                                    .SelectMany(rk => pivotMap.TryGetValue(rk, out var cm)
                                        ? cm.Values.SelectMany(v => v)
                                        : Enumerable.Empty<object[]>())
                                    .ToList(),
                            RelativeTo.GrandTotals => args.AllValuesInOrder,
                            RelativeTo.ParentColTotal =>
                                ResolveSubtotalParentColGroupValues(entry, colLeaves, groupRowKeys, pivotMap),
                            RelativeTo.ParentRowTotal =>
                                entry.IsSubtotal
                                    ? entry.GroupLeaves
                                        .SelectMany(gl =>
                                            pivotMap.Values.SelectMany(cm =>
                                                cm.TryGetValue(gl.PivotKey, out var cv)
                                                    ? cv : Enumerable.Empty<object[]>()))
                                        .ToList()
                                    : pivotMap.Values
                                        .SelectMany(cm => cm.TryGetValue(entry.Leaf.PivotKey, out var cv2)
                                            ? cv2 : Enumerable.Empty<object[]>())
                                        .ToList(),
                            _ => args.AllValuesInOrder
                        };
                    }

                    var val = Aggregate(f, cellVals, context, relativeToVals);
                    result.SetValue(outputRow, col, val ?? string.Empty);
                }
                else
                {
                    result.SetValue(outputRow, col, string.Empty);
                }
                col++;
            }

            if (showColTotal)
            {
                var allGroupVals = groupRowKeys
                    .SelectMany(rk => pivotMap.TryGetValue(rk, out var cm)
                        ? cm.Values.SelectMany(v => v)
                        : Enumerable.Empty<object[]>())
                    .ToList();

                List<object[]> relVals = null;
                if (f.EtaFunction?.Name == "PERCENTOF")
                {
                    var effectiveRelativeTo = args.RelativeTo == RelativeTo.ColumnTotals
                        ? RelativeTo.ParentRowTotal
                        : args.RelativeTo;
                    relVals = effectiveRelativeTo == RelativeTo.GrandTotals || effectiveRelativeTo == RelativeTo.ParentRowTotal
                        ? args.AllValuesInOrder
                        : allGroupVals;
                }

                var totalVal = allGroupVals.Count > 0 ? Aggregate(f, allGroupVals, context, relVals) : null;
                result.SetValue(outputRow, grandTotalCol, totalVal ?? string.Empty);
            }
        }


        private void WriteGrandTotalRow(
            InMemoryRange result,
            int startRow,
            List<ColEntry> colEntries,
            List<LeafWithPath> colLeaves,
            Dictionary<string, Dictionary<string, List<object[]>>> pivotMap,
            PivotByArgs args,
            ParsingContext context,
            int nRowKeyCols,
            int functionNameCol,
            int dataColStart,
            int colOffset,
            int grandTotalCol,
            bool showColTotal,
            int colsPerEntry,
            bool isVStack,
            int functionColOffset)
        {
            var functionNames = ResolveFunctionHeaders(args.Functions);
            int nFunctions = args.Functions.Count;
            int rowsPerTotal = isVStack ? nFunctions : 1;
            string rowTotalLabel = Math.Abs(args.RowTotalDepth) > 1 ? "Grand Total" : "Total";

            for (int fi = 0; fi < rowsPerTotal; fi++)
            {
                int r = startRow + fi;
                var f = args.Functions[fi];

                result.SetValue(r, 0, rowTotalLabel);
                for (int c = 1; c < nRowKeyCols + functionColOffset; c++)
                    result.SetValue(r, c, string.Empty);

                if (isVStack)
                    result.SetValue(r, functionNameCol, functionNames[fi]);

                int col = dataColStart + colOffset;
                foreach (var entry in colEntries)
                {
                    if (entry.IsSubtotal)
                    {
                        var groupVals = entry.GroupLeaves
                            .SelectMany(l =>
                            {
                                var ck = l.PivotKey;
                                return pivotMap.Values
                                    .SelectMany(cm => cm.TryGetValue(ck, out var cv)
                                        ? cv : Enumerable.Empty<object[]>());
                            }).ToList();

                        if (isVStack)
                        {
                            object val = groupVals.Count > 0
                                ? Aggregate(f, groupVals, context,
                                    f.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null)
                                : null;
                            result.SetValue(r, col, val ?? string.Empty);
                            col++;
                        }
                        else
                        {
                            foreach (var func in args.Functions)
                            {
                                object val = groupVals.Count > 0
                                    ? Aggregate(func, groupVals, context,
                                        func.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null)
                                    : null;
                                result.SetValue(r, col, val ?? string.Empty);
                                col++;
                            }
                        }
                    }
                    else
                    {
                        string colKey = entry.Leaf.PivotKey;
                        var allValsForCol = pivotMap.Values
                            .SelectMany(cm => cm.TryGetValue(colKey, out var cv)
                                ? cv : Enumerable.Empty<object[]>())
                            .ToList();

                        if (isVStack)
                        {
                            object grandVal = null;
                            if (allValsForCol.Count > 0)
                            {
                                var effectiveRelativeTo = f.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo == RelativeTo.ColumnTotals
                                    ? RelativeTo.ParentRowTotal
                                    : args.RelativeTo;
                                var relativeToVals = f.EtaFunction?.Name == "PERCENTOF"
                                    ? ResolveRelativeToValuesForTotal(effectiveRelativeTo, entry.Leaf, colLeaves, pivotMap, args) ?? args.AllValuesInOrder
                                    : args.AllValuesInOrder;
                                grandVal = Aggregate(f, allValsForCol, context, relativeToVals);
                            }
                            result.SetValue(r, col, grandVal ?? string.Empty);
                            col++;
                        }
                        else
                        {
                            foreach (var func in args.Functions)
                            {
                                object grandVal = null;
                                if (allValsForCol.Count > 0)
                                {
                                    var effectiveRelativeTo = func.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo == RelativeTo.ColumnTotals
                                        ? RelativeTo.ParentRowTotal
                                        : args.RelativeTo;
                                    var relativeToVals = func.EtaFunction?.Name == "PERCENTOF"
                                        ? ResolveRelativeToValuesForTotal(effectiveRelativeTo, entry.Leaf, colLeaves, pivotMap, args) ?? args.AllValuesInOrder
                                        : args.AllValuesInOrder;
                                    grandVal = Aggregate(func, allValsForCol, context, relativeToVals);
                                }
                                result.SetValue(r, col, grandVal ?? string.Empty);
                                col++;
                            }
                        }
                    }
                }

                if (showColTotal)
                {
                    var allVals = args.AllValuesInOrder.Select(v => new object[] { v[0] }).ToList();
                    int totalCol = grandTotalCol;
                    if (isVStack)
                    {
                        var effectiveRelativeTo = f.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo == RelativeTo.ColumnTotals
                            ? RelativeTo.ParentRowTotal
                            : args.RelativeTo;
                        var cornerRelVals = f.EtaFunction?.Name == "PERCENTOF"
                            ? (effectiveRelativeTo == RelativeTo.ParentRowTotal || effectiveRelativeTo == RelativeTo.GrandTotals
                                ? allVals : args.AllValuesInOrder)
                            : args.AllValuesInOrder;
                        var cornerVal = Aggregate(f, allVals, context, cornerRelVals);
                        result.SetValue(r, totalCol, cornerVal ?? string.Empty);
                    }
                    else
                    {
                        foreach (var func in args.Functions)
                        {
                            var effectiveRelativeTo = func.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo == RelativeTo.ColumnTotals
                                ? RelativeTo.ParentRowTotal
                                : args.RelativeTo;
                            var cornerRelVals = func.EtaFunction?.Name == "PERCENTOF"
                                ? (effectiveRelativeTo == RelativeTo.ParentRowTotal || effectiveRelativeTo == RelativeTo.GrandTotals
                                    ? allVals : args.AllValuesInOrder)
                                : args.AllValuesInOrder;
                            var cornerVal = Aggregate(func, allVals, context, cornerRelVals);
                            result.SetValue(r, totalCol, cornerVal ?? string.Empty);
                            totalCol++;
                        }
                    }
                }

                if (!isVStack) break;
            }
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
                    result.Add(new LeafWithPath(level, path, MakePivotKey(path)));
                else
                    result.AddRange(CollectLeavesWithPath(level.Children, path));
            }
            return result;
        }
    }
}