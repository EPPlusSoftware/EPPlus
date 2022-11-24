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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Utilities;
using IndexFunc = OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Index;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers
{
    public class RpnFunctionCompilerFactory
    {
        private readonly Dictionary<Type, RpnFunctionCompiler> _specialCompilers = new Dictionary<Type, RpnFunctionCompiler>();
        private readonly ParsingContext _context;
        public RpnFunctionCompilerFactory(FunctionRepository repository, ParsingContext context)
        {
            Require.That(context).Named("context").IsNotNull();
            _context = context;
            _specialCompilers.Add(typeof(If), new RpnIfFunctionCompiler(repository.GetFunction("if"), context));
            _specialCompilers.Add(typeof(SumIf), new RpnSumIfCompiler(repository.GetFunction("sumif"), context));
            _specialCompilers.Add(typeof(CountIfs), new RpnCountIfsCompiler(repository.GetFunction("countifs"), context));
            _specialCompilers.Add(typeof(IfError), new RpnIfErrorFunctionCompiler(repository.GetFunction("iferror"), context));
            _specialCompilers.Add(typeof(IfNa), new RpnIfNaFunctionCompiler(repository.GetFunction("ifna"), context));
            _specialCompilers.Add(typeof(Row), new RpnIgnoreCircularRefLookupCompiler(repository.GetFunction("row"), context));
            _specialCompilers.Add(typeof(Rows), new RpnIgnoreCircularRefLookupCompiler(repository.GetFunction("rows"), context));
            _specialCompilers.Add(typeof(Column), new RpnIgnoreCircularRefLookupCompiler(repository.GetFunction("column"), context));
            _specialCompilers.Add(typeof(Columns), new RpnIgnoreCircularRefLookupCompiler(repository.GetFunction("columns"), context));
            _specialCompilers.Add(typeof(IndexFunc), new RpnIgnoreCircularRefLookupCompiler(repository.GetFunction("index"), context));
            foreach (var key in repository.CustomCompilers.Keys)
            {
              _specialCompilers.Add(key, repository.RpnCustomCompilers[key]);
            }
        }

        private RpnFunctionCompiler GetCompilerByType(ExcelFunction function)
        {
            var funcType = function.GetType();
            if (_specialCompilers.ContainsKey(funcType))
            {
                return _specialCompilers[funcType];
            }
            else if (function.IsLookupFuction) return new RpnLookupFunctionCompiler(function, _context);
            else if (function.IsErrorHandlingFunction) return new RpnErrorHandlingFunctionCompiler(function, _context);
            return new RpnDefaultCompiler(function, _context);
        }
        internal virtual RpnFunctionCompiler Create(ExcelFunction function)
        { 
            return GetCompilerByType(function);
        }
    }
}
