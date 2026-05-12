using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.GroupingFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.PivotBy
{
    internal partial class PivotBy
    {
        private void BuildPivotData(
      PivotByArgs args,
      ParsingContext context,
      out List<LeafWithPath> rowLeaves,
      out List<LeafWithPath> colLeaves,
      out Dictionary<string, Dictionary<string, List<object[]>>> pivotMap)
        {
            var resolvedHeaders = ResolveHeaders(args.Headers, args.Values);
            bool hasHeaders = resolvedHeaders == FieldHeaders.YesAndShow
                           || resolvedHeaders == FieldHeaders.YesAndDontShow;
            int startRow = hasHeaders ? 1 : 0;

            int nRowKeyCols = args.RowFields.Size.NumberOfCols;
            int nColKeyCols = args.ColFields.Size.NumberOfCols;
            int nValCols = args.Values.Size.NumberOfCols;

            var rowLeafDict = new Dictionary<string, LeafWithPath>(StringComparer.OrdinalIgnoreCase);
            var rowLeafOrder = new List<string>();
            var colLeafDict = new Dictionary<string, LeafWithPath>(StringComparer.OrdinalIgnoreCase);
            var colLeafOrder = new List<string>();
            pivotMap = new Dictionary<string, Dictionary<string, List<object[]>>>(StringComparer.OrdinalIgnoreCase);

            int nRows = args.RowFields.Size.NumberOfRows;

            for (int r = startRow; r < nRows; r++)
            {
                if (args.FilterArray != null)
                {
                    var fv = args.FilterArray.GetOffset(r, 0);
                    if (fv is bool b && !b) continue;
                    if (fv is int i && i == 0) continue;
                }

                var rowKeyParts = new object[nRowKeyCols];
                for (int c = 0; c < nRowKeyCols; c++)
                    rowKeyParts[c] = args.RowFields.GetOffset(r, c);

                var colKeyParts = new object[nColKeyCols];
                for (int c = 0; c < nColKeyCols; c++)
                    colKeyParts[c] = args.ColFields.GetOffset(r, c);

                var vals = new object[nValCols];
                for (int c = 0; c < nValCols; c++)
                    vals[c] = args.Values.GetOffset(r, c);

                // Radlöv
                string rowKey = MakePivotKey(rowKeyParts);
                if (!rowLeafDict.ContainsKey(rowKey))
                {
                    var leaf = new GroupLevel { Key = rowKeyParts[rowKeyParts.Length - 1] };
                    rowLeafDict[rowKey] = new LeafWithPath(leaf, rowKeyParts);
                    rowLeafOrder.Add(rowKey);
                }
                var rowLeafEntry = rowLeafDict[rowKey];
                var existingRow = rowLeafEntry.Leaf.Rows.FirstOrDefault(rw => MakePivotKey(rw.KeyParts) == rowKey);
                if (existingRow == null)
                {
                    existingRow = new GroupRow { KeyParts = rowKeyParts };
                    rowLeafEntry.Leaf.Rows.Add(existingRow);
                }
                existingRow.Values.Add(vals);

                // Kolumnlöv
                string colKey = MakePivotKey(colKeyParts);
                if (!colLeafDict.ContainsKey(colKey))
                {
                    var leaf = new GroupLevel { Key = colKeyParts[colKeyParts.Length - 1] };
                    colLeafDict[colKey] = new LeafWithPath(leaf, colKeyParts);
                    colLeafOrder.Add(colKey);
                }

                // Pivotkartan
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

            rowLeaves = rowLeafOrder.Select(k => rowLeafDict[k]).ToList();
            colLeaves = colLeafOrder.Select(k => colLeafDict[k]).ToList();

            // Aggregera SubtotalValues per radlöv (används för rad-totals)
            foreach (var rl in rowLeaves)
                AggregateLeaf(rl.Leaf, args, context);
        }

        private void AggregateLeaf(GroupLevel leaf, PivotByArgs args, ParsingContext context)
        {
            var allVals = leaf.Rows.SelectMany(r => r.Values).ToList();
            leaf.SubtotalValues = args.Functions.Select(f =>
            {
                int nValCols = allVals[0].Length;
                var result = new object[nValCols];
                for (int col = 0; col < nValCols; col++)
                {
                    var colValues = allVals.Select(v => new object[] { v[col] }).ToList();
                    result[col] = Aggregate(f, colValues, context,
                        f.EtaFunction?.Name == "PERCENTOF" ? args.AllValuesInOrder : null);
                }
                return result;
            }).ToList();
            leaf.SubtotalValue = leaf.SubtotalValues[0][0];
        }


        private List<LeafWithPath> ApplyRowSort(List<LeafWithPath> rowLeaves, PivotByArgs args) =>
    ApplyLeafSort(rowLeaves, args.RowSortOrders);

        private List<LeafWithPath> ApplyColSort(List<LeafWithPath> colLeaves, PivotByArgs args) =>
            ApplyLeafSort(colLeaves, args.ColSortOrders);

        private List<LeafWithPath> ApplyLeafSort(List<LeafWithPath> leaves, int[] sortOrders)
        {
            if (sortOrders == null || sortOrders.All(s => s == 0)) return leaves;

            IOrderedEnumerable<LeafWithPath> ordered = null;
            foreach (var sortOrder in sortOrders)
            {
                if (sortOrder == 0) continue;
                bool desc = sortOrder < 0;
                int col = Math.Abs(sortOrder) - 1;
                var capturedCol = col;

                Func<LeafWithPath, object> keySelector = lp =>
                    capturedCol < lp.Path.Length ? lp.Path[capturedCol] : null;

                if (ordered == null)
                    ordered = desc
                        ? leaves.OrderByDescending(keySelector, _comparer)
                        : leaves.OrderBy(keySelector, _comparer);
                else
                    ordered = desc
                        ? ordered.ThenByDescending(keySelector, _comparer)
                        : ordered.ThenBy(keySelector, _comparer);
            }

            // Bryt oavgjort med efterföljande nyckeldelar
            int maxDepth = leaves.Max(l => l.Path.Length);
            int sortedDepth = sortOrders.Max(s => Math.Abs(s));
            for (int col = sortedDepth; col < maxDepth; col++)
            {
                var capturedCol = col;
                Func<LeafWithPath, object> keySelector = lp =>
                    capturedCol < lp.Path.Length ? lp.Path[capturedCol] : null;
                ordered = ordered.ThenBy(keySelector, _comparer);
            }

            return ordered?.ToList() ?? leaves;
        }
    }
}
