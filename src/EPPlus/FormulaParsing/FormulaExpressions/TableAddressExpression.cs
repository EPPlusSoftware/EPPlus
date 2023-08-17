using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Diagnostics;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    [DebuggerDisplay("TableAddressExpression: {_addressInfo}")]
    internal class TableAddressExpression : Expression
    {
        readonly FormulaTableAddress _addressInfo;
        private bool _negate;

        public TableAddressExpression(FormulaTableAddress addressInfo, ParsingContext ctx) : base(ctx)
        {
            _addressInfo = addressInfo;
        }
        internal override ExpressionType ExpressionType => ExpressionType.TableAddress;

        public override CompileResult Compile()
        {
            if (_cachedCompileResult == null)
            {
                if (_addressInfo.FromRow < 1)
                {
                    _cachedCompileResult = new CompileResult(eErrorType.Ref);
                }

                var ri = Context.ExcelDataProvider.GetRange(_addressInfo);
                if (ri.IsMulti)
                {
                    _cachedCompileResult = new AddressCompileResult(ri, DataType.ExcelRange, _addressInfo);
                }
                else
                {
                    _cachedCompileResult = CompileResultFactory.Create(ri.GetOffset(0, 0), _addressInfo);
                }
            }
            return _cachedCompileResult;
        }

        public override void Negate()
        {
            _negate = !_negate;
        }
        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.CanCompile;
        public override FormulaRangeAddress GetAddress() 
        { 
            return _addressInfo.Clone();
        }
    }
}
