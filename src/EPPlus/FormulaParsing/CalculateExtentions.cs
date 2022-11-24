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

using System.Threading;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.Core.CellStore;

namespace OfficeOpenXml
{
    /// <summary>
    /// Extentions methods for formula calculation.
    /// </summary>
    public static class CalculationExtension
    {
        /// <summary>
        /// Calculate all formulas in the current workbook
        /// </summary>
        /// <param name="workbook">The workbook</param>
        public static void Calculate(this ExcelWorkbook workbook)
        {
            Calculate(workbook, new ExcelCalculationOption(){AllowCircularReferences=false});
        }

        /// <summary>
        /// Calculate all formulas in the current workbook
        /// </summary>
        /// <param name="workbook">The workbook to calculate</param>
        /// <param name="configHandler">Configuration handler</param>
        /// <example>
        /// <code>
        /// workbook.Calculate(opt => opt.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
        /// </code>
        /// </example>
        public static void Calculate(this ExcelWorkbook workbook, Action<ExcelCalculationOption> configHandler)
        {
            var option = new ExcelCalculationOption();
            configHandler.Invoke(option);
            Calculate(workbook, option);
        }


        /// <summary>
        /// Calculate all formulas in the current workbook
        /// </summary>
        /// <param name="workbook">The workbook</param>
        /// <param name="options">Calculation options</param>
        public static void Calculate(this ExcelWorkbook workbook, ExcelCalculationOption options)
        {
            Init(workbook);

            var dc = DependencyChainFactory.Create(workbook, options);
            var filterInfo = new FilterInfo(workbook);
            workbook.FormulaParser.InitNewCalc(filterInfo);
            if (workbook.FormulaParser.Logger != null)
            {
                var msg = string.Format("Starting... number of cells to parse: {0}", dc.list.Count);
                workbook.FormulaParser.Logger.Log(msg);
            }

            CalcChain(workbook, workbook.FormulaParser, dc, options);
        }
        /// <summary>
        /// Calculate all formulas in the current worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet</param>
        public static void Calculate(this ExcelWorksheet worksheet)
        {
            Calculate(worksheet, new ExcelCalculationOption());
        }

        /// <summary>
        /// Calculate all formulas in the current range
        /// </summary>
        /// <param name="worksheet">The worksheet to calculate</param>
        /// <param name="configHandler">Configuration handler</param>
        /// <example>
        /// <code>
        /// sheet.Calculate(opt => opt.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
        /// </code>
        /// </example>
        public static void Calculate(this ExcelWorksheet worksheet, Action<ExcelCalculationOption> configHandler)
        {
            var option = new ExcelCalculationOption();
            configHandler.Invoke(option);
            Calculate(worksheet, option);
        }

        /// <summary>
        /// Calculate all formulas in the current worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet</param>
        /// <param name="options">Calculation options</param>
        public static void Calculate(this ExcelWorksheet worksheet, ExcelCalculationOption options)
        {
            Init(worksheet.Workbook);       
            var dc = DependencyChainFactory.Create(worksheet, options);
            var parser = worksheet.Workbook.FormulaParser;
            var filterInfo = new FilterInfo(worksheet.Workbook);
            parser.InitNewCalc(filterInfo);
            if (parser.Logger != null)
            {
                var msg = string.Format("Starting... number of cells to parse: {0}", dc.list.Count);
                parser.Logger.Log(msg);
            }
            CalcChain(worksheet.Workbook, parser, dc, options);
        }
        /// <summary>
        /// Calculate all formulas in the current range
        /// </summary>
        /// <param name="range">The range</param>
        public static void Calculate(this ExcelRangeBase range)
        {
            Calculate(range, new ExcelCalculationOption());
        }

        /// <summary>
        /// Calculate all formulas in the current range
        /// </summary>
        /// <param name="range">The range to calculate</param>
        /// <param name="configHandler">Configuration handler</param>
        /// <example>
        /// <code>
        /// sheet.Cells["A1:A3"].Calculate(opt => opt.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
        /// </code>
        /// </example>
        public static void Calculate(this ExcelRangeBase range, Action<ExcelCalculationOption> configHandler)
        {
            var option = new ExcelCalculationOption();
            configHandler.Invoke(option);
            Calculate(range, option);
        }

        /// <summary>
        /// Calculate all formulas in the current range
        /// </summary>
        /// <param name="range">The range</param>
        /// <param name="options">Calculation options</param>
        public static void Calculate(this ExcelRangeBase range, ExcelCalculationOption options)
        {
            Init(range._workbook);
            var parser = range._workbook.FormulaParser;
            var filterInfo = new FilterInfo(range._workbook);
            parser.InitNewCalc(filterInfo);
            var dc = DependencyChainFactory.Create(range, options);
            CalcChain(range._workbook, parser, dc, options);
        }

        /// <summary>
        /// Calculate all formulas in the current range
        /// </summary>
        /// <param name="worksheet">The worksheet</param>
        /// <param name="Formula">The formula to be calculated</param>
        /// <returns>The result of the formula calculation</returns>
        public static object Calculate(this ExcelWorksheet worksheet, string Formula)
        {
            return Calculate(worksheet, Formula, new ExcelCalculationOption());
        }
        /// <summary>
        /// Calculate all formulas in the current range
        /// </summary>
        /// <param name="worksheet">The worksheet</param>
        /// <param name="Formula">The formula to be calculated</param>
        /// <param name="options">Calculation options</param>
        /// <returns>The result of the formula calculation</returns>
        public static object Calculate(this ExcelWorksheet worksheet, string Formula, ExcelCalculationOption options)
        {
            try
            {
                worksheet.CheckSheetTypeAndNotDisposed();
                if(string.IsNullOrEmpty(Formula.Trim())) return null;
                Init(worksheet.Workbook);
                var parser = worksheet.Workbook.FormulaParser;
                var filterInfo = new FilterInfo(worksheet.Workbook);
                parser.InitNewCalc(filterInfo);
                if (Formula[0] == '=') Formula = Formula.Substring(1); //Remove any starting equal sign
                var dc = DependencyChainFactory.Create(worksheet, Formula, options);
                var f = dc.list[0];
                dc.CalcOrder.RemoveAt(dc.CalcOrder.Count - 1);

                CalcChain(worksheet.Workbook, parser, dc, options);

                return parser.ParseCell(f.Tokens, worksheet.Name, -1, -1);
            }
            catch (Exception ex)
            {
                return new ExcelErrorValueException(ex.Message, ExcelErrorValue.Create(eErrorType.Value));
            }
        }
        private static void CalcChain(ExcelWorkbook wb, FormulaParser parser, DependencyChain dc, ExcelCalculationOption options)
        {
            wb.FormulaParser.Configure(config =>
            {
                config.AllowCircularReferences = options.AllowCircularReferences;
                config.PrecisionAndRoundingStrategy = options.PrecisionAndRoundingStrategy;
            });
            var debug = parser.Logger != null;
            foreach (var ix in dc.CalcOrder)
            {
                var item = dc.list[ix];
                try
                {
                    object v;
                    if (item.wsIndex >= 0 && item.wsIndex < wb.Worksheets.Count)
                    {
                        var ws = wb.Worksheets._worksheets[item.wsIndex];
                        v = parser.ParseCell(item.Tokens, ws == null ? "" : ws.Name, item.Row, item.Column);
                    }
                    else
                    {
                        if(item.Column == 0 && item.Row >= 0 && item.Row < wb.Names.Count)
                        {
                            v = parser.ParseCell(item.Tokens, null, item.Row, item.Column);
                        }
                        else
                        {
                            v = ExcelErrorValue.Create(eErrorType.Ref);
                        }
                    }
                    
                    SetValue(wb, item, v);
                    if (debug)
                    {
                        parser.Logger.LogCellCounted();
                    }
                    if(ix % 1000 == 0)
                    {
                        Thread.Sleep(0);
                    }
                }
                catch(Exception e)
                {
                    if(e is CircularReferenceException)
                    {
                        throw;
                    }
                    else
                    {
                        var error = ExcelErrorValue.Parse(ExcelErrorValue.Values.Value);
                        SetValue(wb, item, error);
                    }
                }
            }
        }
        internal static void Init(ExcelWorkbook workbook)
        {
            workbook._formulaTokens = new CellStore<IList<Token>>();
            foreach (var ws in workbook.Worksheets)
            {
                if (!(ws is ExcelChartsheet))
                {
                    if (ws._formulaTokens != null)
                    {
                        ws._formulaTokens.Dispose();
                    }
                    ws._formulaTokens = new CellStore<IList<Token>>();
                }
            }
        }
        private static void SetValue(ExcelWorkbook workbook, FormulaCell item, object v)
        {
            if (item.Column == 0)
            {
                if (item.wsIndex < 0)
                {
                    workbook.Names[item.Row].NameValue = v;
                }
                else
                {
                    var sh = workbook.Worksheets._worksheets[item.wsIndex];
                    sh.Names[item.Row].NameValue = v;
                }
            }
            else
            {
                var sheet = workbook.Worksheets._worksheets[item.wsIndex];
                sheet.SetValueInner(item.Row, item.Column, v);
            }
        }
    }
}
