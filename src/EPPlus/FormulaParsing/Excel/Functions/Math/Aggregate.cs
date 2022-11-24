/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
         Category = ExcelFunctionCategory.Statistical,
         EPPlusVersion = "5.5",
         IntroducedInExcelVersion = "2010",
         Description = "Performs a specified calculation (e.g. the sum, product, average, etc.) for a list or database, with the option to ignore hidden rows and error values")]
    internal class Aggregate : ExcelFunction
    {
        public override void BeforeInvoke(ParsingContext context)
        {
            base.BeforeInvoke(context);
            var cellId = context.ExcelDataProvider.GetCellId(context.Scopes.Current.Address.WorksheetName, context.Scopes.Current.Address.FromRow, context.Scopes.Current.Address.FromCol);
            if (!context.SubtotalAddresses.Contains(cellId))
            {
                context.SubtotalAddresses.Add(cellId);
            }
        }
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var funcNum = ArgToInt(arguments, 0);
            var nToSkip = IsNumeric(arguments.ElementAt(1).Value) ? 2 : 1;  
            var options = nToSkip == 1 ? 0 : ArgToInt(arguments, 1);

            if (options < 0 || options > 7) return CreateResult(eErrorType.Value);

            if(IgnoreNestedSubtotalAndAggregate(options))
            {
                context.Scopes.Current.IsSubtotal = true;
                var cellId = context.ExcelDataProvider.GetCellId(context.Scopes.Current.Address.WorksheetName, context.Scopes.Current.Address.FromRow, context.Scopes.Current.Address.FromCol);
                if(!context.SubtotalAddresses.Contains(cellId))
                {
                    context.SubtotalAddresses.Add(cellId);
                }
            }

            CompileResult result = null;
            switch(funcNum)
            {
                case 1:
                    var f1 = new Average()
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f1.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 2:
                    var f2 = new Count()
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f2.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 3:
                    var f3 = new CountA
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f3.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 4:
                    var f4 = new Max 
                    { 
                        IgnoreHiddenValues = IgnoreHidden(options), 
                        IgnoreErrors = IgnoreErrors(options) 
                    };
                    result = f4.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 5:
                    var f5 = new Min
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f5.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 6:
                    var f6 = new Product
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f6.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 7:
                    var f7 = new StdevDotS
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f7.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 8:
                    var f8 = new StdevDotP
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f8.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 9:
                    var f9 = new Sum
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f9.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 10:
                    VarDotS f10 = new VarDotS
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f10.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 11:
                    var f11 = new VarDotP
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f11.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 12:
                    var f12 = new Median
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f12.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 13:
                    var f13 = new ModeSngl
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f13.Execute(arguments.Skip(nToSkip), context);
                    break;
                case 14:
                    var f14 = new Large
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    var a141 = arguments.ElementAt(nToSkip);
                    var a142 = arguments.ElementAt(nToSkip + 1);
                    result = f14.Execute(new List<FunctionArgument> { a141, a142 }, context);
                    break;
                case 15:
                    var f15 = new Small
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f15.Execute(new List<FunctionArgument> { arguments.ElementAt(nToSkip), arguments.ElementAt(nToSkip + 1) }, context);
                    break;
                case 16:
                    var f16 = new PercentileInc
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f16.Execute(new List<FunctionArgument> { arguments.ElementAt(nToSkip), arguments.ElementAt(nToSkip + 1) }, context);
                    break;
                case 17:
                    var f17 = new QuartileInc
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f17.Execute(new List<FunctionArgument> { arguments.ElementAt(nToSkip), arguments.ElementAt(nToSkip + 1) }, context);
                    break;
                case 18:
                    var f18 = new PercentileExc
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f18.Execute(new List<FunctionArgument> { arguments.ElementAt(nToSkip), arguments.ElementAt(nToSkip + 1) }, context);
                    break;
                case 19:
                    var f19 = new QuartileExc
                    {
                        IgnoreHiddenValues = IgnoreHidden(options),
                        IgnoreErrors = IgnoreErrors(options)
                    };
                    result = f19.Execute(new List<FunctionArgument> { arguments.ElementAt(nToSkip), arguments.ElementAt(nToSkip + 1) }, context);
                    break;
                default:
                    result = CreateResult(eErrorType.Value);
                    break;
            }
            result.IsResultOfSubtotal = IgnoreNestedSubtotalAndAggregate(options);
            return result;
        }

        private bool IgnoreHidden(int options)
        {
            return options == 1 || options == 3 || options == 5 || options == 7;
        }

        private bool IgnoreErrors(int options)
        {
            return options == 2 || options == 3 || options == 6 || options == 7;
        }

        private bool IgnoreNestedSubtotalAndAggregate(int options)
        {
            return options == 0 || options == 1 || options == 2 || options == 3;
        }
    }
}
