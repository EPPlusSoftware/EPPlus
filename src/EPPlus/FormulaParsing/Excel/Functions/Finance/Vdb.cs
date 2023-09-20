/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  16/08/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/

using OfficeOpenXml.Encryption;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
    Category = ExcelFunctionCategory.Financial,
    EPPlusVersion = "7.0",
    Description = "Calculates the depreciation of an asset over a specific period. Calculates the depreciation with either" +
        "double declining method or straight line method.")]
    internal class Vdb : ExcelFunction
    {
        public override int ArgumentMinLength => 5;
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(0, 1, 2, 3, 4, 5);
        }

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var cost = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);

            var salvage = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CompileResult.GetErrorResult(e2.Type);

            var life = ArgToDecimal(arguments, 2, out ExcelErrorValue e3);
            if (e3 != null) return CompileResult.GetErrorResult(e3.Type);

            var startPeriod = ArgToDecimal(arguments, 3, out ExcelErrorValue e4);
            if (e4 != null) return CompileResult.GetErrorResult(e4.Type);

            var endPeriod = ArgToDecimal(arguments, 4, out ExcelErrorValue e5);
            if (e5 != null) return CompileResult.GetErrorResult(e5.Type);

            var factor = 2d;
            var noSwitch = false;
            if (arguments.Count > 5)
            {
                factor = ArgToDecimal(arguments, 5, out ExcelErrorValue e6);
                if (e6 != null) return CompileResult.GetErrorResult(e6.Type);
            }
            if (arguments.Count > 6) noSwitch = ArgToBool(arguments, 6);
            if (cost < 0) return CreateResult(eErrorType.Num);
            if (salvage < 0) return CreateResult(eErrorType.Num);
            if (life <= 0) return CreateResult(eErrorType.Num);
            if (startPeriod > life) return CreateResult(eErrorType.Num);
            if (endPeriod > life) return CreateResult(eErrorType.Num);
            if (startPeriod > endPeriod) return CreateResult(eErrorType.Num);
            if (factor < 0) return CreateResult(eErrorType.Num);

            var assetDepreciation = (noSwitch) ? DepreciationOverPeriod(cost, salvage, life, endPeriod, factor, false) -
                                                 DepreciationOverPeriod(cost, salvage, life, startPeriod, factor, false)
                                                 :
                                                 DepreciationOverPeriod(cost, salvage, life, endPeriod, factor, true) -
                                                 DepreciationOverPeriod(cost, salvage, life, startPeriod, factor, true);

            return CreateResult(assetDepreciation, DataType.Decimal);
        }

        public static double DepreciationOverPeriod(double cost, double salvage, double life, double period, double factor, bool straightLine)
        {
            double GetDoubleDecliningBalance(double currentDep, double per)
            {
                double fractionOfPeriod = period - Math.Truncate(period);

                double doubleDecliningDep = Math.Min(((cost - currentDep) * (factor / life)), (cost - salvage - currentDep));
                double straightLineDep = (cost - currentDep - salvage) / (life - per);

                bool performSwitch = straightLine && doubleDecliningDep < straightLineDep; 
                double periodDep = performSwitch ? straightLineDep : doubleDecliningDep;
                double cumulatedDep = currentDep + periodDep;

                if ((int)period == 0d)
                    return cumulatedDep * fractionOfPeriod;
                else if ((int)per == (int)period - 1)
                {
                    double doubleDecliningBalanceNextPeriod = 
                    Math.Min(((cost - cumulatedDep) * (factor / life)), (cost - salvage - cumulatedDep));

                    double straightLineNextPeriod = (cost - cumulatedDep - salvage) / (life - (per + 1d)); 
                    bool isSlnNextPeriod = straightLine && doubleDecliningBalanceNextPeriod < straightLineNextPeriod;

                    double deprNextPeriod = isSlnNextPeriod ?
                        (period == life ? 0d : straightLineNextPeriod) :
                        doubleDecliningBalanceNextPeriod;

                    return cumulatedDep + deprNextPeriod * fractionOfPeriod;
                }
                else //Gets accelerated depreciation up until end period.
                {
                    return GetDoubleDecliningBalance(cumulatedDep, per + 1d);
                }
            }

            return GetDoubleDecliningBalance(0d, 0d); //starting at 0 depreciation at period 0
        }

    }
}
