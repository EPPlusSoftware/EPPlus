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
using System.Linq;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class NamedValueExpression : Expression
    {
        internal short _externalReferenceIx;
        internal int _worksheetIx;
        internal INameInfo _name;        
        int _negate = 0; //0 if no negation is performed. -1 or 1 if the value should be negated. In this case the value is converted to a double and negated. If the value is non numeric #VALUE is returned.
        internal NamedValueExpression(string name, ParsingContext parsingContext, short externalReferenceIx, int worksheetIx) : base(parsingContext)
        {
            _externalReferenceIx = externalReferenceIx;
            _worksheetIx = worksheetIx;
            _name = Context.ExcelDataProvider.GetName(_externalReferenceIx, worksheetIx, name);
        }
        private NamedValueExpression(INameInfo nameInfo, ParsingContext parsingContext, short externalReferenceIx, int worksheetIx, int negate) : base(parsingContext)
        {
            _externalReferenceIx = externalReferenceIx;
            _worksheetIx = worksheetIx;
            _name = nameInfo;
            _negate = negate;
        }

        internal override ExpressionType ExpressionType => ExpressionType.NameValue;
        public override CompileResult Compile()
        {
            if (_name == null) return new CompileResult(ExcelErrorValue.Create(eErrorType.Name), DataType.ExcelError);

            var value = _name.GetValue(Context.CurrentCell);
            if (value==null)
            {
                return new CompileResult(null, DataType.Empty);
            }

            if (value is IRangeInfo)
            {
                var range = (IRangeInfo)value;
                if (range.GetNCells() > 1)
                {
                    if(_negate == -1)
                    {
                        range = CreateNegatedRange(range);
                    }
                    return new AddressCompileResult(range, DataType.ExcelRange, range.Address);
                }
                else
                {                    
                    if (range.IsEmpty)
                    {
                        return new AddressCompileResult(null, DataType.Empty, range.Address);
                    }
                    var v = range.GetOffset(0, 0);
                    return GetNegatedValue(v, range.Address);
                }
            }
            else
            {
                
                return GetNegatedValue(value, null);
            }            
        }
        public bool IsRelative 
        { 
            get 
            { 
                return _name.IsRelative;
            } 
        }
        private CompileResult GetNegatedValue(object value, FormulaRangeAddress address)
        {
            if (_negate == 0)
            {
                return CompileResultFactory.Create(value, address);
            }
            else if(value is ExcelErrorValue e)
            {
                return CompileResult.GetErrorResult(e.Type);
            }
            else
            {
                var d = ConvertUtil.GetValueDouble(value, false, true);
                if (double.IsNaN(d))
                {
                    return CompileResultFactory.Create(ExcelErrorValue.Create(eErrorType.Value), address);
                }
                else
                {
                    return CompileResultFactory.Create(d * _negate);
                }
            }
        }

        private InMemoryRange CreateNegatedRange(IRangeInfo range)
        {
            var resultRange = new InMemoryRange(range.Size);
            for(var row = 0; row < range.Size.NumberOfRows; row++)
            {
                for(var col = 0; col < range.Size.NumberOfCols; col++)
                {
                    var v = range.GetOffset(row, col);
                    var d = ConvertUtil.GetValueDouble(v, false, true);
                    if (double.IsNaN(d))
                    {
                        resultRange.SetValue(row, col, ExcelErrorValue.Create(eErrorType.Value));
                    }
                    else
                    {
                        resultRange.SetValue(row, col, d * -1);
                    }
                }
            }
            return resultRange;
        }

        public override Expression Negate()
        {
            int n;
            if (_negate == 0)
            {
                n = -1;
            }
            else
            {
                n = _negate * -1;
            }
            return new NamedValueExpression(_name, Context, _externalReferenceIx, _worksheetIx, n);
        }
        public override FormulaRangeAddress[] GetAddress()
        {

            if(_name?.Value is IRangeInfo ri) 
            {
                if(_name.IsRelative)
                {
                    return _name.GetRelativeRange(ri, Context.CurrentCell).Addresses;
                }
                else
                {
                    return ri.Addresses;
                }
            }
            return null;
        }
        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.CanCompile;
    }
}
