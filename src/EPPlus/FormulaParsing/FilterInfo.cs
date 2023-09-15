/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/18/2021         EPPlus Software AB       Improved handling of hidden cells for SUBTOTAL and AGGREGATE.
 *************************************************************************************************/
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.Filter;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// This class contains information of the usage of Filters on the worksheets of a workbook.
    /// One area where this information is needed is when running the SUBTOTAL function. If
    /// there is an active filter on the worksheet hidden cells should be ignored even if SUBTOTAL
    /// is called with a single digit func num.
    /// </summary>
    internal class FilterInfo
    {
        public FilterInfo(ExcelWorkbook workbook)
        {
            _workbook = workbook;
            Initialize();
        }

        private readonly ExcelWorkbook _workbook;
        private readonly Dictionary<int, QuadTree<ExcelAutoFilter>> _worksheetFilters = new Dictionary<int, QuadTree<ExcelAutoFilter>>();

        private void Initialize()
        {
            foreach(var worksheet in _workbook.Worksheets)
            {
                if (worksheet.IsChartSheet) continue;
                var endRow = 1;
                var endCol = 1;
                if(worksheet.Dimension != null && worksheet.Dimension.Start != null && worksheet.Dimension.End != null)
                {
                    endRow = worksheet.Dimension.End.Row;
                    endCol = worksheet.Dimension.End.Column;
                }
                var qt = new QuadTree<ExcelAutoFilter>(1, 1, endRow, endCol);
                if(worksheet.AutoFilter != null)
                {
                    var r = new QuadRange(worksheet.AutoFilter.Address);
                    qt.Add(r, worksheet.AutoFilter);
                }
                foreach(var table in worksheet.Tables)
                {
                    if (table.AutoFilter != null && table.AutoFilter.Columns != null && table.AutoFilter.Columns.Count > 0)
                    {
                        var r = new QuadRange(table.Address);
                        qt.Add(r, table.AutoFilter);
                    }
                }
                _worksheetFilters.Add(worksheet.IndexInList, qt);
                //if(worksheet.AutoFilter != null && worksheet.AutoFilter.Columns != null && worksheet.AutoFilter.Columns.Count > 0)
                //{
                //    _worksheetFilters.Add(worksheet.IndexInList);
                //    continue;
                //}
                //foreach(var table in worksheet.Tables)
                //{                    
                //    if(table.AutoFilter != null && table.AutoFilter.Columns != null && table.AutoFilter.Columns.Count > 0)
                //    {
                //        if(!_worksheetFilters.Contains(worksheet.IndexInList))
                //        {
                //            _worksheetFilters.Add(worksheet.IndexInList);
                //            continue;
                //        }
                //    }
                //}
            }
        }

        /// <summary>
        /// Returns true if there is an Autofilter with at least one column on the requested worksheet.
        /// </summary>
        /// <param name="wsIx">Worksheet index</param>
        /// <param name="cell">Cell to check</param>
        /// <returns></returns>
        public bool CellIsCoveredByFilter(int wsIx, ICellInfo cell)
        {
            if (!_worksheetFilters.ContainsKey(wsIx) || cell.Address == null) return false;
            var qt = _worksheetFilters[wsIx];
            var qr = new QuadRange(cell.Row, cell.Column, cell.Row, cell.Column);
            var ir = qt.GetIntersectingRanges(qr);
            return ir.Count > 0;
        }

        /// <summary>
        /// Returns true if there is an Autofilter with at least one column on the requested worksheet.
        /// </summary>
        /// <param name="wsIx">Worksheet index</param>
        /// <param name="cell">Cell to check</param>
        /// <returns></returns>
        public bool CellIsCoveredByFilter(int wsIx, FunctionArgument arg)
        {
            if (!_worksheetFilters.ContainsKey(wsIx) || arg.Address == null) return false;
            var qt = _worksheetFilters[wsIx];
            var qr = new QuadRange(arg.Address.FromRow, arg.Address.FromCol, arg.Address.ToRow, arg.Address.ToCol);
            var ir = qt.GetIntersectingRanges(qr);
            return ir.Count > 0;
        }
    }
}
