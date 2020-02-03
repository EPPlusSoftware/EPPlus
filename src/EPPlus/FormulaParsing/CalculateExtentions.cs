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
        /// <param name="workbook">The workbook</param>
        /// <param name="options">Calculation options</param>
        public static void Calculate(this ExcelWorkbook workbook, ExcelCalculationOption options)
        {
            Init(workbook);

            var dc = DependencyChainFactory.Create(workbook, options);
            workbook.FormulaParser.InitNewCalc();
            if (workbook.FormulaParser.Logger != null)
            {
                var msg = string.Format("Starting... number of cells to parse: {0}", dc.list.Count);
                workbook.FormulaParser.Logger.Log(msg);
            }

            CalcChain(workbook, workbook.FormulaParser, dc);
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
        /// Calculate all formulas in the current worksheet
        /// </summary>
        /// <param name="worksheet">The worksheet</param>
        /// <param name="options">Calculation options</param>
        public static void Calculate(this ExcelWorksheet worksheet, ExcelCalculationOption options)
        {
            Init(worksheet.Workbook);
            //worksheet.Workbook._formulaParser = null; TODO:Cant reset. Don't work with userdefined or overrided worksheet functions            
            var dc = DependencyChainFactory.Create(worksheet, options);
            var parser = worksheet.Workbook.FormulaParser;
            parser.InitNewCalc();
            if (parser.Logger != null)
            {
                var msg = string.Format("Starting... number of cells to parse: {0}", dc.list.Count);
                parser.Logger.Log(msg);
            }
            CalcChain(worksheet.Workbook, parser, dc);
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
        /// <param name="range">The range</param>
        /// <param name="options">Calculation options</param>
        public static void Calculate(this ExcelRangeBase range, ExcelCalculationOption options)
        {
            Init(range._workbook);
            var parser = range._workbook.FormulaParser;
            parser.InitNewCalc();
            var dc = DependencyChainFactory.Create(range, options);
            CalcChain(range._workbook, parser, dc);
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
                worksheet.CheckSheetType();
                if(string.IsNullOrEmpty(Formula.Trim())) return null;
                Init(worksheet.Workbook);
                var parser = worksheet.Workbook.FormulaParser;
                parser.InitNewCalc();
                if (Formula[0] == '=') Formula = Formula.Substring(1); //Remove any starting equal sign
                var dc = DependencyChainFactory.Create(worksheet, Formula, options);
                var f = dc.list[0];
                dc.CalcOrder.RemoveAt(dc.CalcOrder.Count - 1);

                CalcChain(worksheet.Workbook, parser, dc);

                return parser.ParseCell(f.Tokens, worksheet.Name, -1, -1);
            }
            catch (Exception ex)
            {
                return new ExcelErrorValueException(ex.Message, ExcelErrorValue.Create(eErrorType.Value));
            }
        }
        private static void CalcChain(ExcelWorkbook wb, FormulaParser parser, DependencyChain dc)
        {
            var debug = parser.Logger != null;
            foreach (var ix in dc.CalcOrder)
            {
                var item = dc.list[ix];
                try
                {
                    var ws = wb.Worksheets.GetBySheetID(item.SheetID);
                    var v = parser.ParseCell(item.Tokens, ws == null ? "" : ws.Name, item.Row, item.Column);
                    SetValue(wb, item, v);
                    if (debug)
                    {
                        parser.Logger.LogCellCounted();
                    }
                    Thread.Sleep(0);
                }
                catch (FormatException fe)
                {
                    throw (fe);
                }
                catch(CircularReferenceException cre)
                {
                    throw cre;
                }
                catch
                {
                    var error = ExcelErrorValue.Parse(ExcelErrorValue.Values.Value);
                    SetValue(wb, item, error);
                }
            }
        }
        private static void Init(ExcelWorkbook workbook)
        {
            workbook._formulaTokens = new CellStore<List<Token>>();;
            foreach (var ws in workbook.Worksheets)
            {
                if (!(ws is ExcelChartsheet))
                {
                    if (ws._formulaTokens != null)
                    {
                        ws._formulaTokens.Dispose();
                    }
                    ws._formulaTokens = new CellStore<List<Token>>();
                }
            }
        }
        private static void SetValue(ExcelWorkbook workbook, FormulaCell item, object v)
        {
            if (item.Column == 0)
            {
                if (item.SheetID <= 0)
                {
                    workbook.Names[item.Row].NameValue = v;
                }
                else
                {
                    var sh = workbook.Worksheets.GetBySheetID(item.SheetID);
                    sh.Names[item.Row].NameValue = v;
                }
            }
            else
            {
                var sheet = workbook.Worksheets.GetBySheetID(item.SheetID);
                sheet.SetValueInner(item.Row, item.Column, v);
            }
        }
    }
}
