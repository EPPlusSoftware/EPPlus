/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/3/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils
{
    internal static class ImplicitIntersectionUtil
    {
        public static CompileResult GetResult(IRangeInfo range, FormulaCellAddress currentCell, ParsingContext context)
        {
            return GetResult(range, currentCell.Row, currentCell.Column, context);
        }
        public static CompileResult GetResult(IRangeInfo range, int ccr, int ccc, ParsingContext context)
        {
            if (range.Address==null) return CompileResult.GetErrorResult(eErrorType.Value);
            var fr = range.Address.FromRow;
            var tr = range.Address.ToRow;
            var fc = range.Address.FromCol;
            var tc = range.Address.ToCol;
            // always return #VALUE if both multiple rows and multiple cols
            if (tr - fr > 0 && tc - fc > 0) return CompileResult.GetErrorResult(eErrorType.Value);

            object result;
            FormulaRangeAddress addr;
            //Single cell, always return the value and the address.
            if (fr == tr && fc == tc) 
            {
                result = range.GetValue(fr, fc);
                addr = new FormulaRangeAddress(context, fr, fc, tr, tc);
                return CompileResultFactory.Create(result, addr);
            }

            // if current cell is outside rows and cols of the range
            // are we outside the allowed area?
            if ((ccr < fr || ccr > tr) && (ccc < fc || ccc > tc)) return CompileResult.GetErrorResult(eErrorType.Value);

            // do implicit intersection

            if(tr - fr > 0)
            {
                if (ccr < fr || ccr > tr) return CompileResult.GetErrorResult(eErrorType.Value);
                // use row of the current cell
                result = range.GetValue(ccr, tc);
                addr = new FormulaRangeAddress(context, ccr, tc, ccr, tc);
            }
            // horizontal direction
            else
            {
                if (ccc < fc || ccc > tc) return CompileResult.GetErrorResult(eErrorType.Value);
                // use col of the current cell
                result = range.GetValue(tr, ccc);
                addr = new FormulaRangeAddress(context, tr, ccc, tr, ccc);
            }

            return CompileResultFactory.Create(result, addr);
        }
    }
}
