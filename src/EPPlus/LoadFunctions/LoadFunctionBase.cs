/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/16/2020         EPPlus Software AB       EPPlus 5.2.1
 *************************************************************************************************/
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    /// <summary>
    /// Base class for ExcelRangeBase.LoadFrom[...] functions
    /// </summary>
    internal abstract class LoadFunctionBase
    {
        public LoadFunctionBase(ExcelRangeBase range, LoadFunctionFunctionParamsBase parameters)
        {
            Range = range;
            PrintHeaders = parameters.PrintHeaders;
            TableStyle = parameters.TableStyle;
            TableName = parameters.TableName?.Trim();
        }

        /// <summary>
        /// The range to which the data should be loaded
        /// </summary>
        protected ExcelRangeBase Range { get; }

        /// <summary>
        /// If true a header row will be printed above the data
        /// </summary>
        protected bool PrintHeaders { get; }

        /// <summary>
        /// If value is other than TableStyles.None the data will be added to a table in the worksheet.
        /// </summary>
        protected TableStyles? TableStyle { get; set; }
        protected string TableName { get; set; }

        protected bool ShowFirstColumn { get; set; }

        protected bool ShowLastColumn { get; set; }

        protected bool ShowTotal { get; set; }

        /// <summary>
        /// Returns how many rows there are in the range (header row not included)
        /// </summary>
        /// <returns></returns>
        protected abstract int GetNumberOfRows();

        /// <summary>
        /// Returns how many columns there are in the range
        /// </summary>
        /// <returns></returns>
        protected abstract int GetNumberOfColumns();

        protected virtual void PostProcessTable(ExcelTable table, ExcelRangeBase range)
        {

        }

        protected abstract void LoadInternal(object[,] values, out Dictionary<int, FormulaCell> formulaCells, out Dictionary<int, string> columnFormats);

        /// <summary>
        /// Loads the data into the worksheet
        /// </summary>
        /// <returns></returns>
        internal ExcelRangeBase Load()
        {
            var nRows = PrintHeaders ? GetNumberOfRows() + 1 : GetNumberOfRows();
            var nCols = GetNumberOfColumns();
            var values = new object[nRows, nCols];

            //if(Range.Worksheet._values.Capacity < values.Length)
            //{
            //    Range.Worksheet._values.Capacity = values.Length;
            //}

            LoadInternal(values, out Dictionary<int, FormulaCell> formulaCells, out Dictionary<int, string> columnFormats);
            var ws = Range.Worksheet;
            if(formulaCells != null && formulaCells.Keys.Count > 0)
            {
                SetValuesAndFormulas(nRows, nCols, values, formulaCells, ws);
            }
            else
            {
                ws.SetRangeValueInner(Range._fromRow, Range._fromCol, Range._fromRow + nRows - 1, Range._fromCol + nCols - 1, values, true);
            }


            //Must have at least 1 row, if header is shown
            if (nRows == 1 && PrintHeaders)
            {
                nRows++;
            }
            // set number formats
            foreach (var col in columnFormats.Keys)
            {
                ws.Cells[Range._fromRow, Range._fromCol + col, Range._fromRow + nRows - 1, Range._fromCol + col].Style.Numberformat.Format = columnFormats[col];
            }

            if(nRows==0)
            {
                return null;
            }

            var r = ws.Cells[Range._fromRow, Range._fromCol, Range._fromRow + nRows - 1, Range._fromCol + nCols - 1];

            if (TableStyle.HasValue)
            {
                var tbl = ws.Tables.Add(r, TableName);
                tbl.ShowHeader = PrintHeaders;
                tbl.TableStyle = TableStyle.Value;
                tbl.ShowFirstColumn = ShowFirstColumn;
                tbl.ShowLastColumn = ShowLastColumn;
                tbl.ShowTotal = ShowTotal;
                PostProcessTable(tbl, r);
            }
            return r;
        }

        private void SetValuesAndFormulas(int nRows, int nCols, object[,] values, Dictionary<int, FormulaCell> formulaCells, ExcelWorksheet ws)
        {
            for (var col = 0; col < nCols; col++)
            {
                if (formulaCells.ContainsKey(col))
                {
                    var row = 0;
                    if (PrintHeaders)
                    {
                        var header = values[0, col];
                        ws.SetValue(Range._fromRow, Range._fromCol + col, header);
                        row++;
                    }
                    var columnFormula = formulaCells[col];
                    var fromRow = Range._fromRow + row;
                    var rangeCol = Range._fromCol + col;
                    var toRow = Range._fromRow + nRows - 1;
                    var formulaRange = ws.Cells[fromRow, rangeCol, toRow, rangeCol];
                    if (!string.IsNullOrEmpty(columnFormula.Formula))
                    {
                        formulaRange.Formula = columnFormula.Formula;
                    }
                    else
                    {
                        formulaRange.FormulaR1C1 = columnFormula.FormulaR1C1;
                    }
                }
                else
                {
                    object[,] columnValues = new object[nRows, 1];
                    for (var ix = 0; ix < nRows; ix++)
                    {
                        var item = values[ix, col];
                        columnValues[ix, 0] = item;
                    }
                    var fromRow = Range._fromRow;
                    var rangeCol = Range._fromCol + col;
                    var toRow = Range._fromRow + nRows - 1;
                    ws.SetRangeValueInner(fromRow, rangeCol, toRow, rangeCol, columnValues, true);
                }

            }
        }
    }
}
