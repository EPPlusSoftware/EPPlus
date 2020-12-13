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
        protected TableStyles TableStyle { get; set; }

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

        protected virtual void PostProcessTable(ExcelTable table)
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
            LoadInternal(values, out Dictionary<int, FormulaCell> formulaCells, out Dictionary<int,string> columnFormats);
            var ws = Range.Worksheet;
            ws.SetRangeValueInner(Range._fromRow, Range._fromCol, Range._fromRow + nRows - 1, Range._fromCol + nCols - 1, values);
            
            //Must have at least 1 row, if header is shown
            if (nRows == 1 && PrintHeaders)
            {
                nRows++;
            }
            // set number formats
            foreach(var col in columnFormats.Keys)
            {
                ws.Cells[Range._fromRow, Range._fromCol + col, Range._fromRow + nRows - 1, Range._fromCol + col].Style.Numberformat.Format = columnFormats[col];
            }
            // set formulas
            foreach(var col in formulaCells.Keys)
            {
                var row = PrintHeaders ? 0 : -1;
                while (++row < nRows)
                {
                    var formulaCell = formulaCells[col];
                    if(!string.IsNullOrEmpty(formulaCell.Formula))
                    {
                        var formula = formulaCell.Formula.Replace("{row}", (Range._fromRow + row).ToString());
                        ws.SetFormula(Range._fromRow + row, Range._fromCol + col, formula);
                    }
                    else if(!string.IsNullOrEmpty(formulaCell.FormulaR1C1))
                    {
                        ws.Cells[Range._fromRow + row, Range._fromCol + col].FormulaR1C1 = formulaCell.FormulaR1C1;
                    }
                }
            }

            var r = ws.Cells[Range._fromRow, Range._fromCol, Range._fromRow + nRows - 1, Range._fromCol + nCols - 1];

            if (TableStyle != TableStyles.None)
            {
                var tbl = ws.Tables.Add(r, "");
                tbl.ShowHeader = PrintHeaders;
                tbl.TableStyle = TableStyle;
                PostProcessTable(tbl);
            }

            return r;
        }
    }
}
