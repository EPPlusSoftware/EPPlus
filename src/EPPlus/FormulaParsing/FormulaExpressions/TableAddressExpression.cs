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
            if (_addressInfo.FromRow < 1)
            {
                return new CompileResult(eErrorType.Ref);
            }

            var ri = Context.ExcelDataProvider.GetRange(_addressInfo);
            if (ri.GetNCells() > 1)
            {
                return new AddressCompileResult(ri, DataType.ExcelRange, _addressInfo);
            }
            else
            {
                return CompileResultFactory.Create(ri.GetOffset(0, 0), _addressInfo);
            }
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
