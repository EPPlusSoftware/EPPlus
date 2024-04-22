﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2024         EPPlus Software AB       Initial release EPPlus 7.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "7.2",
        Description = "Splits a string into substrings",
        SupportsArrays = true)]
    internal class TextSplit : ExcelFunctionTextBase
    {
        public override int ArgumentMinLength => 2;
        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var range = ArgToRangeInfo(arguments, 0);
            var text = string.Empty;
            if (range == null)
            {
                text = ArgToString(arguments, 0);
            }
            string colDelimiter = ArgDelimiterCollectionToString(arguments, 1, out CompileResult result);
            if (result != null) return result;
            string rowDelimiter = string.Empty;
            var ignoreEmpty = "0";
            var matchMode = "0";
            var padWith = "#N/A";
            if (arguments.Count > 2 && arguments[2].Value != null)
            {
                rowDelimiter = ArgDelimiterCollectionToString(arguments, 2, out result);
                if (result != null) return result;
            }
            if (arguments.Count > 3 && arguments[3].Value != null)
            {
                ignoreEmpty = ArgToString(arguments, 3).ToUpper();
            }
            if (arguments.Count > 4 && arguments[4].Value != null)
            {
                matchMode = ArgToString(arguments, 4);
                if (matchMode == "1")
                {
                    colDelimiter += colDelimiter.ToLower() + colDelimiter.ToUpper();
                    rowDelimiter += rowDelimiter.ToLower() + rowDelimiter.ToUpper();
                }
            }
            if (arguments.Count > 5 && arguments[5].Value != null)
            {
                padWith = ArgToString(arguments, 5);
            }

            var rows = new string[] { text };
            if ( !string.IsNullOrEmpty( rowDelimiter))
            {
                rows = (ignoreEmpty == "1" || ignoreEmpty == "TRUE") ? text.Split(rowDelimiter.ToCharArray(), StringSplitOptions.RemoveEmptyEntries) : text.Split(rowDelimiter.ToCharArray());
            }
            var cols = (ignoreEmpty == "1" || ignoreEmpty == "TRUE") ? text.Split(colDelimiter.ToCharArray(), StringSplitOptions.RemoveEmptyEntries) : text.Split(colDelimiter.ToCharArray());
            var returnRange = new InMemoryRange(rows.Length, (short)cols.Length);
            for (var row = 0; row < rows.Length; row++)
            {
                string[] rowCols = (ignoreEmpty == "1" || ignoreEmpty=="TRUE") ? rows[row].Split(colDelimiter.ToCharArray(), StringSplitOptions.RemoveEmptyEntries) : rows[row].Split(colDelimiter.ToCharArray());
                for (var col = 0; col < cols.Length; col++)
                {
                    if (rowCols.Length < cols.Length && col >= rowCols.Length)
                    {
                        if (padWith == "#N/A")
                        {
                            returnRange.SetValue(row, col, CompileResult.GetErrorResult(eErrorType.NA));
                        }
                        else
                        {
                            returnRange.SetValue(row, col, padWith);
                        }
                    }
                    else
                    {
                        returnRange.SetValue(row, col, rowCols[col]);
                    }
                }
            }
            return CreateDynamicArrayResult(returnRange, DataType.ExcelRange);
        }
    }
}
