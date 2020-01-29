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
            //var result = _parsingContext.Parser.Parse(value.ToString());

            if (name == null)
            {
                throw (new Exceptions.ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Name)));
            }
            if (name.Value==null)
            {
                return null;
            }
            if (name.Value is ExcelDataProvider.IRangeInfo)
            {
                var range = (ExcelDataProvider.IRangeInfo)name.Value;
                if (range.IsMulti)
                {
                    return new CompileResult(name.Value, DataType.Enumerable);
                }
                else
                {
                    if (range.IsEmpty)
                    {
                        return null;
                    }
                    var factory = new CompileResultFactory();
                    return factory.Create(range.First().Value);
                }
            }
            else
            {                
                var factory = new CompileResultFactory();
                return factory.Create(name.Value);
            }

            
            
            //return new CompileResultFactory().Create(result);
        }
    }
}
