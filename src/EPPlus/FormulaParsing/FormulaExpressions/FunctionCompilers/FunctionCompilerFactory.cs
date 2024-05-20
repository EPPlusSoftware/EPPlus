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
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Utilities;
using IndexFunc = OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Index;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers
{
    internal class FunctionCompilerFactory
    {
        //private readonly Dictionary<Type, FunctionCompiler> _specialCompilers = new Dictionary<Type, FunctionCompiler>();
        
        public FunctionCompilerFactory(FunctionRepository repository)
        {
            //foreach (var key in repository.CustomCompilers.Keys)
            //{
            //  _specialCompilers.Add(key, repository.CustomCompilers[key]);
            //}
        }

        private FunctionCompiler GetCompilerByType(ExcelFunction function, ParsingContext context)
        {
            var funcType = function.GetType();
            if (
                function.ArrayBehaviour == ExcelFunctionArrayBehaviour.Custom
                ||
                function.ArrayBehaviour == ExcelFunctionArrayBehaviour.FirstArgCouldBeARange)
            {
                return new CustomArrayBehaviourCompiler(function, context);
            }
            return new DefaultCompiler(function);
        }
        internal virtual FunctionCompiler Create(ExcelFunction function, ParsingContext context)
        { 
            return GetCompilerByType(function, context);
        }
    }
}
