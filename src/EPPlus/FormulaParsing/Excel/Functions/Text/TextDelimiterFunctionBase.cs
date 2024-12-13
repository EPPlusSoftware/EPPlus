/*************************************************************************************************
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
using OfficeOpenXml.FormulaParsing.ExcelUtilities.TextUtils;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    internal abstract class TextDelimiterFunctionBase : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override string NamespacePrefix => "_xlfn.";

        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        protected List<string> GetDelimiters(IList<FunctionArgument> arguments)
        {
            var arg2 = arguments[1];
            var delimiters = new List<string>();
            if (arg2.IsExcelRange)
            {
                foreach (var delimiter in arg2.ValueAsRangeInfo)
                {
                    delimiters.Add(delimiter.Value.ToString());
                }
            }
            else
            {
                var del = ArgToString(arguments, 1);
                delimiters.Add(del);
            }
            return delimiters;
        }

        protected int GetInstanceNum(IList<FunctionArgument> arguments)
        {
            var instanceNum = 1;
            if (arguments.Count > 2)
            {
                instanceNum = ArgToInt(arguments, 2, RoundingMethod.Convert);
                if (instanceNum == 0)
                {
                    instanceNum = 1;
                }
            }
            return instanceNum;
        }

        protected int GetMatchMode(IList<FunctionArgument> arguments, out ExcelErrorValue e)
        {
            var matchMode = 0;
            e = null;
            if (arguments.Count > 3)
            {
                matchMode = ArgToInt(arguments, 3, out ExcelErrorValue e1);
                if (e1 != null) e = e1;
                if (matchMode < 0 || matchMode > 1)
                {
                    e = ExcelErrorValue.Create(eErrorType.Value);
                }
            }
            return matchMode;
        }

        protected int GetMatchEnd(IList<FunctionArgument> arguments, out ExcelErrorValue e)
        {
            var matchEnd = 0;
            e = null;
            if (arguments.Count > 4)
            {
                matchEnd = ArgToInt(arguments, 4, RoundingMethod.Convert);
                if (matchEnd < 0 || matchEnd > 1)
                {
                    e = ExcelErrorValue.Create(eErrorType.Value);
                }
            }
            return matchEnd;
        }

        protected string GetIfNotFound(IList<FunctionArgument> arguments)
        {
            if (arguments.Count > 5)
            {
                return ArgToString(arguments, 5);
            }
            return null;
        }

        protected CompileResult ProcessText(bool getBefore, string text, List<string> delimiters, int instanceNum, int matchMode, int matchEnd, string ifNotFound)
        {
            var splitUtil = new TextSplitUtil(text);
            var minPos = -1;
            var currentTextBefore = string.Empty;
            var emptyDelimiterExists = false;
            foreach (var delimiter in delimiters)
            {
                if (string.IsNullOrEmpty(delimiter))
                {
                    emptyDelimiterExists = true;
                    continue;
                }
                if(getBefore)
                {
                    var tb = splitUtil.GetTextBefore(delimiters, instanceNum, matchMode == 1, matchEnd == 1, out bool isOutOfRange, out int? matchIndex);
                    if (!isOutOfRange && matchIndex.HasValue)
                    {
                        if (minPos == -1 || matchIndex.Value < minPos)
                        {
                            minPos = matchIndex.Value;
                            currentTextBefore = tb;
                        }
                    }
                }
                else
                {
                    var tb = splitUtil.GetTextAfter(delimiters, instanceNum, matchMode == 1, matchEnd == 1, out bool isOutOfRange, out int? matchIndex);
                    if (!isOutOfRange && matchIndex.HasValue)
                    {
                        if (minPos == -1 || matchIndex.Value < minPos)
                        {
                            minPos = matchIndex.Value;
                            currentTextBefore = tb;
                        }
                    }
                }
            }
            if (minPos == -1)
            {
                if (emptyDelimiterExists)
                {
                    return new CompileResult(string.Empty, DataType.String);
                }
                if (ifNotFound == null)
                {
                    return CompileResult.GetErrorResult(eErrorType.NA);
                }
                currentTextBefore = ifNotFound;
            }
            return new CompileResult(currentTextBefore, DataType.String);
        }

        protected CompileResult ProcessTextBefore(string text, List<string> delimiters, int instanceNum, int matchMode, int matchEnd, string ifNotFound)
        {
            return ProcessText(true, text, delimiters, instanceNum, matchMode, matchEnd, ifNotFound);
        }

        protected CompileResult ProcessTextAfter(string text, List<string> delimiters, int instanceNum, int matchMode, int matchEnd, string ifNotFound)
        {
            return ProcessText(false, text, delimiters, instanceNum, matchMode, matchEnd, ifNotFound);
        }
    }
}
