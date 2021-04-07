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
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class NamedValueExpression : AtomicExpression
    {
        public NamedValueExpression(string expression, ParsingContext parsingContext)
            : base(expression)
        {
            _parsingContext = parsingContext;
        }

        private readonly ParsingContext _parsingContext;

        public override CompileResult Compile()
        {
            var c = this._parsingContext.Scopes.Current;
            var name = _parsingContext.ExcelDataProvider.GetName(c.Address.Worksheet, ExpressionString);
            
            var cache = _parsingContext.AddressCache;
            var cacheId = cache.GetNewId();
            
            if(!cache.Add(cacheId, ExpressionString))
            {
                throw new InvalidOperationException("Catastropic error occurred, address caching failed");
            }

            if (name == null)
            {
                // check if there is a table with the name
                var table = _parsingContext.ExcelDataProvider.GetExcelTable(ExpressionString);
                if(table != null)
                {
                    var ri = new RangeInfo(table.WorkSheet, table.Address);
                    cache.Add(cacheId, ri.Address.FullAddress);
                    return new CompileResult(ri, DataType.Enumerable, cacheId);
                }
                return new CompileResult(eErrorType.Name);
            }
            if (name.Value==null)
            {
                return new CompileResult(null, DataType.Empty, cacheId);
            }
            if (name.Value is ExcelDataProvider.IRangeInfo)
            {
                var range = (ExcelDataProvider.IRangeInfo)name.Value;
                if (range.IsMulti)
                {
                    cache.Add(cacheId, range.Address.FullAddress);
                    return new CompileResult(name.Value, DataType.Enumerable, cacheId);
                }
                else
                {
                    if (range.IsEmpty)
                    {
                        return new CompileResult(null, DataType.Empty, cacheId);
                    }
                    var factory = new CompileResultFactory();
                    return factory.Create(range.First().Value, cacheId);
                }
            }
            else
            {                
                var factory = new CompileResultFactory();
                return factory.Create(name.Value, cacheId);
            }

            
            
            //return new CompileResultFactory().Create(result);
        }
    }
}
