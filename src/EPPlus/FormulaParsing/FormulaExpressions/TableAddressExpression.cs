using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Diagnostics;
using System.Collections.Generic;

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
            if (_addressInfo.FromRow < 1 || _addressInfo.FromCol < 1 ||
                _addressInfo.ToRow < 1 || _addressInfo.ToCol < 1)
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

        public override Expression Negate()
        {
            _negate = !_negate;
            return this;
        }
        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.CanCompile;
        public override FormulaRangeAddress[] GetAddress() 
        { 
            return [_addressInfo.Clone()];
        }
        internal override Expression CloneWithOffset(int row, int col)
        {
            if (row == 0 && col == 0)
            {
                return this;
            }
            var ai = new FormulaRangeAddress(Context)
            {
                ExternalReferenceIx = _addressInfo.ExternalReferenceIx,
                WorksheetIx = _addressInfo.WorksheetIx,
                FromRow = (_addressInfo.FixedFlag & FixedFlag.FromRowFixed) == FixedFlag.FromRowFixed ? _addressInfo.FromRow : _addressInfo.FromRow + row,
                ToRow = (_addressInfo.FixedFlag & FixedFlag.ToRowFixed) == FixedFlag.ToRowFixed ? _addressInfo.ToRow : _addressInfo.ToRow + row,
                FromCol = (_addressInfo.FixedFlag & FixedFlag.FromColFixed) == FixedFlag.FromColFixed ? _addressInfo.FromCol : _addressInfo.FromCol + col,
                ToCol = (_addressInfo.FixedFlag & FixedFlag.ToColFixed) == FixedFlag.ToColFixed ? _addressInfo.ToCol : _addressInfo.ToCol + col,
            };
            return new RangeExpression(ai)
            {
                Status = Status,
                Operator = Operator
            };
        }

    }
}
