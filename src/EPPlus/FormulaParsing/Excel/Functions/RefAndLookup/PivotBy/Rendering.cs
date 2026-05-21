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
        private List<object[]> ResolveRelativeToValues(
            RelativeTo relativeTo,
            LeafWithPath colLeaf,
            List<LeafWithPath> colLeaves,
            string rowKey,
            Dictionary<string, Dictionary<string, List<object[]>>> pivotMap,
            PivotByArgs args)
        {
            switch (relativeTo)
            {
                case RelativeTo.RowTotals: // 1
                    {
                        if (!pivotMap.TryGetValue(rowKey, out var colMap))
                            return null;

                        return colMap.Values
                            .SelectMany(vals => vals)
                            .ToList();
                    }
                case RelativeTo.GrandTotals: // 2
                    {
                        // Nämnare = alla värden i hela datasetet
                        return args.AllValuesInOrder;
                    }
                case RelativeTo.ParentColTotal:
                    {
                        if (!pivotMap.TryGetValue(rowKey, out var colMap))
                            return null;

                        // En kolumnnivå = ingen riktig föräldragrupp finns.
                        // Faller tillbaka på RowTotals: nämnare = radens totalsumma över ALLA kolumner.
                        if (colLeaf.Path.Length <= 1)
                        {
                            return colMap.Values
                                .SelectMany(vals => vals)
                                .ToList();
                        }

                        // Flernivå: nämnare = radens värden inom förälderns kolumngrupp
                        var parentKey = colLeaf.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty;
                        var siblingLeaves = colLeaves
                            .Where(l => (l.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty) == parentKey)
                            .ToList();

                        return siblingLeaves
                            .Where(leaf => colMap.ContainsKey(MakePivotKey(leaf.Path)))
                            .SelectMany(leaf => colMap[MakePivotKey(leaf.Path)])
                            .ToList();
                    }
                case RelativeTo.ParentRowTotal: // 4
                    {
                        // Nämnare = alla raders värden för denna kolumn (kolumnens totalsumma)
                        string colKey = MakePivotKey(colLeaf.Path);
                        return pivotMap.Values
                            .SelectMany(cm => cm.TryGetValue(colKey, out var cv)
                                ? cv
                                : Enumerable.Empty<object[]>())
                            .ToList();
                    }
                default:
                    return null;
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
                case RelativeTo.RowTotals: // 1
                    {
                        return pivotMap.Values
                            .SelectMany(colMap => colMap.Values.SelectMany(v => v))
                            .ToList();
                    }
                case RelativeTo.GrandTotals: // 2
                    {
                        return args.AllValuesInOrder;
                    }
                case RelativeTo.ParentColTotal:
                    {
                        // En kolumnnivå = faller tillbaka på GrandTotals som nämnare
                        if (colLeaf.Path.Length <= 1)
                        {
                            return pivotMap.Values
                                .SelectMany(colMap => colMap.Values.SelectMany(v => v))
                                .ToList();
                        }

                        var parentKey = colLeaf.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty;
                        var siblingLeaves = colLeaves
                            .Where(l => (l.Path[0]?.ToString()?.ToLowerInvariant() ?? string.Empty) == parentKey)
                            .ToList();

                        return pivotMap.Values
                            .SelectMany(cm => siblingLeaves
                                .Where(leaf => cm.ContainsKey(MakePivotKey(leaf.Path)))
                                .SelectMany(leaf => cm[MakePivotKey(leaf.Path)]))
                            .ToList();
                    }
                case RelativeTo.ParentRowTotal: // 4
                    {
                        // Total-raden: kolumn / kolumn = 1
                        string colKey = MakePivotKey(colLeaf.Path);
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
                foreach (var leaf in groupLeaves)
                    colEntries.Add(new ColEntry { IsSubtotal = false, Leaf = leaf });
                if (showColSubtotals)
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

            // --- Fältnamnrad ---
            if (showFieldHeaders)
            {
                for (int i = 0; i < nRowKeyCols + functionColOffset; i++)
                    result.SetValue(0, i, string.Empty);
                for (int i = 0; i < args.ColFields.Size.NumberOfCols; i++)
                    result.SetValue(0, dataColStart + colOffset + i, args.ColFields.GetOffset(0, i));
            }

            // --- Rubrikrader ---
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
                        result.SetValue(outputLevel, col, f == 0 ? val : string.Empty);
                        col++;
                    }
                }

                if (showColTotal)
                {
                    string colTotalLabel = Math.Abs(args.ColTotalDepth) > 1 ? "Grand Total" : "Total";
                    for (int f = 0; f < colsPerEntry; f++)
                    {
                        result.SetValue(outputLevel, col, f == 0 && level == 0 ? colTotalLabel : string.Empty);
                        col++;
                    }
                }
            }

            // --- HSTACK: Funktionsnamnsrad ---
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
                    foreach (var name in functionNames)
                    {
                        result.SetValue(functionHeaderRow, col, name);
                        col++;
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
                    for (int fc = 0; fc < colsPerEntry; fc++)
                    {
                        result.SetValue(headerDataRow, col, headerValue);
                        col++;
                    }
                }
            }

            // --- Grand total-rad överst ---
            if (rowTotalAtTop && showRowTotal)
                WriteGrandTotalRow(result, fieldHeaderRows + nColKeyRows + functionHeaderRows + headerDataRows,
                                   colEntries, colLeaves, pivotMap, args, context,
                                   nRowKeyCols, functionNameCol, dataColStart, colOffset,
                                   grandTotalCol, showColTotal, colsPerEntry, isVStack, functionColOffset);

            // --- Datarader ---
            var functionNames2 = ResolveFunctionHeaders(args.Functions);
            int currentOutputRow = dataRowStart;

            foreach (var rowGroup in rowGroups)
            {
                var groupLeaves = rowGroup.ToList();

                foreach (var rowLeafEntry in groupLeaves)
                {
                    var rowPath = rowLeafEntry.Path;
                    var rowLeaf = rowLeafEntry.Leaf;
                    string rowKey = MakePivotKey(rowPath);

                    for (int fi = 0; fi < nFunctions; fi++)
                    {
                        int outputRow = currentOutputRow + fi;
                        var f = args.Functions[fi];

                        // Fyll alla radnyckelkolumner – även de som saknas i rowPath
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
                                        var ck = MakePivotKey(l.Path);
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
                                string colKey = MakePivotKey(entry.Leaf.Path);

                                if (isHStack)
                                {
                                    foreach (var func in args.Functions)
                                    {
                                        object aggregated = null;
                                        if (pivotMap.TryGetValue(rowKey, out var colMap) &&
                                            colMap.TryGetValue(colKey, out var cellVals))
                                        {
                                            var effectiveRelativeTo = func.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo == RelativeTo.ColumnTotals
                                                ? RelativeTo.ParentRowTotal
                                                : args.RelativeTo;
                                            var relativeToVals = func.EtaFunction?.Name == "PERCENTOF"
                                                ? ResolveRelativeToValues(effectiveRelativeTo, entry.Leaf, colLeaves, rowKey, pivotMap, args) ?? args.AllValuesInOrder
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
                                        var effectiveRelativeTo = f.EtaFunction?.Name == "PERCENTOF" && args.RelativeTo == RelativeTo.ColumnTotals
                                            ? RelativeTo.ParentRowTotal
                                            : args.RelativeTo;
                                        var relativeToVals = f.EtaFunction?.Name == "PERCENTOF"
                                            ? ResolveRelativeToValues(effectiveRelativeTo, entry.Leaf, colLeaves, rowKey, pivotMap, args) ?? args.AllValuesInOrder
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
                            if (isHStack)
                            {
                                foreach (var func in args.Functions)
                                {
                                    var relVals = args.AllValuesInOrder;
                                    if (func.EtaFunction?.Name == "PERCENTOF")
                                    {
                                        var effectiveRelativeTo = args.RelativeTo == RelativeTo.ColumnTotals
                                            ? RelativeTo.ParentRowTotal
                                            : args.RelativeTo;
                                        relVals = effectiveRelativeTo == RelativeTo.GrandTotals || effectiveRelativeTo == RelativeTo.ParentRowTotal
                                            ? args.AllValuesInOrder
                                            : rowAllVals;
                                    }
                                    var totalVal = Aggregate(func, rowAllVals, context, relVals);
                                    result.SetValue(outputRow, col, totalVal ?? string.Empty);
                                    col++;
                                }
                            }
                            else
                            {
                                var relVals = args.AllValuesInOrder;
                                if (f.EtaFunction?.Name == "PERCENTOF")
                                {
                                    var effectiveRelativeTo = args.RelativeTo == RelativeTo.ColumnTotals
                                        ? RelativeTo.ParentRowTotal
                                        : args.RelativeTo;
                                    relVals = effectiveRelativeTo == RelativeTo.GrandTotals || effectiveRelativeTo == RelativeTo.ParentRowTotal
                                        ? args.AllValuesInOrder
                                        : rowAllVals;
                                }
                                var totalVal = Aggregate(f, rowAllVals, context, relVals);
                                result.SetValue(outputRow, col, totalVal ?? string.Empty);
                                col++;
                            }
                        }

                        if (isHStack) break;
                    }

                    currentOutputRow += rowsPerLeaf;
                }

                // --- Radsubtotalrad för gruppen ---
                if (showRowSubtotals)
                {
                    var groupKeyParts = groupLeaves[0].Path.Take(rowSubtotalDepth - 1).ToArray();
                    for (int fi = 0; fi < rowsPerLeaf; fi++)
                    {
                        WriteRowSubtotalRow(
                            result, currentOutputRow + fi,
                            groupKeyParts, groupLeaves,
                            colEntries, colLeaves,
                            pivotMap, args, context,
                            nRowKeyCols, functionNameCol,
                            dataColStart, colOffset,
                            showColTotal, isVStack,
                            args.Functions[fi], functionNames2[fi]);
                    }
                    currentOutputRow += rowsPerLeaf;
                }
            }

            // --- Grand total-rad nederst ---
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
             bool showColTotal,
             bool isVStack,
             LambdaCalculator f,
             string functionName)
        {
            // Radnycklar: känd prefix + tomt för resten
            for (int k = 0; k < nRowKeyCols; k++)
                result.SetValue(outputRow, k, k < groupKeyParts.Length ? groupKeyParts[k] : string.Empty);

            if (isVStack)
                result.SetValue(outputRow, functionNameCol, functionName);

            var groupRowKeys = groupLeaves.Select(l => MakePivotKey(l.Path)).ToList();

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
                                .SelectMany(gl => cm.TryGetValue(MakePivotKey(gl.Path), out var cv)
                                    ? cv : Enumerable.Empty<object[]>());
                        }).ToList();
                }
                else
                {
                    string colKey = MakePivotKey(entry.Leaf.Path);
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
                                groupRowKeys
                                    .SelectMany(rk => pivotMap.TryGetValue(rk, out var cm)
                                        ? cm.Values.SelectMany(v => v)
                                        : Enumerable.Empty<object[]>())
                                    .ToList(),
                            RelativeTo.ParentRowTotal =>
                                entry.IsSubtotal
                                    ? entry.GroupLeaves
                                        .SelectMany(gl =>
                                            pivotMap.Values.SelectMany(cm =>
                                                cm.TryGetValue(MakePivotKey(gl.Path), out var cv)
                                                    ? cv : Enumerable.Empty<object[]>()))
                                        .ToList()
                                    : pivotMap.Values
                                        .SelectMany(cm => cm.TryGetValue(MakePivotKey(entry.Leaf.Path), out var cv2)
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
                result.SetValue(outputRow, col, totalVal ?? string.Empty);
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
                                var ck = MakePivotKey(l.Path);
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
                        string colKey = MakePivotKey(entry.Leaf.Path);
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
                        result.SetValue(r, col, cornerVal ?? string.Empty);
                        col++;
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
                            result.SetValue(r, col, cornerVal ?? string.Empty);
                            col++;
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
                    result.Add(new LeafWithPath(level, path));
                else
                    result.AddRange(CollectLeavesWithPath(level.Children, path));
            }
            return result;
        }
    }
}
