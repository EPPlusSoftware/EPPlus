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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.XPath;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the number of cells (of a supplied range), that satisfy a set of given criteria",
        IntroducedInExcelVersion = "2007")]
    internal class CountIfs : MultipleRangeCriteriasFunction
    {
        public override int ArgumentMinLength => 2;

        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;
        public override FunctionParameterInformation GetParameterInfo(int argumentIndex)
        {
            if(argumentIndex % 2 == 1)
            {
                return FunctionParameterInformation.IgnoreErrorInPreExecute;
            }
            return FunctionParameterInformation.Normal;
        }

        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.IgnoreNumberOfArgsFromStart = 1;
            config.ArrayArgInterval = 2;
        }
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var argRanges = new List<RangeOrValue>();
            var criterias = new List<object>();
            for (var ix = 0; ix < 30; ix +=2)
            {
                if (arguments.Count <= ix) break;
                var arg = arguments[ix];
                if (arg.DataType == DataType.ExcelError) continue;
                var rangeInfo = arg.ValueAsRangeInfo;
                if(rangeInfo == null && arg.Address !=null)
                {
                    var wsIx = arg.Address.WorksheetIx < 0 ? context.CurrentCell.WorksheetIx : arg.Address.WorksheetIx;
                    rangeInfo = context.ExcelDataProvider.GetRange(wsIx, arg.Address.FromRow, arg.Address.FromCol);
                    argRanges.Add(new RangeOrValue { Range = rangeInfo });
                }
                else if(rangeInfo != null)
                {
                    argRanges.Add(new RangeOrValue { Range = rangeInfo});
                }
                else
                {
                    argRanges.Add(new RangeOrValue { Value = arg.Value });
                }
                criterias.Add(arguments[ix + 1].ValueFirst);
            }
            IEnumerable<int> matchIndexes = GetMatchIndexes(argRanges[0], criterias[0], context, false);
            var enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();
            for (var ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
            {
                var indexes = GetMatchIndexes(argRanges[ix], criterias[ix], context, false);
                matchIndexes = matchIndexes.Intersect(indexes);
            }
            
            return CreateResult((double)matchIndexes.Count(), DataType.Integer);
        }
    }
}
