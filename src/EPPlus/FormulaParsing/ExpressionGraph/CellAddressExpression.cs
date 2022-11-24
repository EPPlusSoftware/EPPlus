using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Diagnostics;
using Operators = OfficeOpenXml.FormulaParsing.Excel.Operators.Operators;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    [DebuggerDisplay("CellAddressExpression: {ExpressionString}")]
    internal class CellAddressExpression : ExpressionWithParent
    {
        FormulaRangeAddress _addressInfo;
        bool _negate;
        public CellAddressExpression(Token token, ParsingContext ctx, ref FormulaAddressBase addressInfo) : base(token.Value, ctx)
        {
            if(addressInfo== null)
            {
                _addressInfo = new FormulaRangeAddress();
            }
            else
            {
                _addressInfo = new FormulaRangeAddress() { ExternalReferenceIx= addressInfo.ExternalReferenceIx, WorksheetIx=addressInfo.WorksheetIx };
            }
            addressInfo = _addressInfo;
            _negate = token.IsNegated;
        }
        public override bool IsGroupedExpression => false;

        public bool HasCircularReference { get; internal set; }

        internal override ExpressionType ExpressionType => ExpressionType.CellAddress;

        public override CompileResult Compile()
        {
            if (_result == null)
            {
                ExcelCellBase.GetRowColFromAddress(ExpressionString, out int row, out int col, out bool fixedRow, out bool fixedCol);
                //Range
                _addressInfo.FromRow = _addressInfo.ToRow = row;
                _addressInfo.FromCol = _addressInfo.ToCol = col;
                _addressInfo.FixedFlag = fixedRow ? FixedFlag.FromRowFixed | FixedFlag.ToRowFixed : 0;
                _addressInfo.FixedFlag |= fixedCol ? FixedFlag.FromColFixed | FixedFlag.ToColFixed : 0;
                _addressInfo.WorksheetIx = _addressInfo.WorksheetIx == short.MinValue ? (short)Context.CurrentCell.WorksheetIx : _addressInfo.WorksheetIx;

                if ((Operator != null && Operator.Operator == Operators.Colon)/* || (Prev != null && Prev.Operator.Operator == Operators.Colon)*/)
                {

                    _result = new AddressCompileResult(_addressInfo, DataType.ExcelCellAddress, _addressInfo);
                }
                else
                {
                    // Single Cell.
                    var wsIx = _addressInfo.WorksheetIx < -1 ? Context.Scopes.Current.Address.WorksheetIx : _addressInfo.WorksheetIx;
                    if (wsIx < 0) return new CompileResult(eErrorType.Ref);
                    _result = CompileResultFactory.Create(Context.Package.Workbook.Worksheets[wsIx].GetValueInner(row, col), 0, _addressInfo);
                    if (_result.IsNumeric && _negate)
                    {
                        _result.Negate();
                    }
                }
            }
            return _result;
        }
        internal override Expression Clone()
        {
            return this;
        }
        internal override Expression Clone(int rowOffset, int colOffset)
        {
            if (_result == null) Compile();
            FormulaAddressBase address=null;
            var exp = new CellAddressExpression(new Token(), Context, ref address);
            var range = (FormulaRangeAddress)address;
            range.ExternalReferenceIx = _addressInfo.ExternalReferenceIx;
            range.WorksheetIx = _addressInfo.WorksheetIx;
            if ((_addressInfo.FixedFlag&FixedFlag.FromRowFixed)==0)
            {
                range.FromRow = _addressInfo.FromRow + rowOffset;
            }
            else
            {
                range.FromRow = _addressInfo.FromRow;
            }

            if ((_addressInfo.FixedFlag & FixedFlag.ToRowFixed) == 0)
            {
                range.ToRow = _addressInfo.ToRow + rowOffset;
            }
            else
            {
                range.ToRow = _addressInfo.ToRow;
            }

            if ((_addressInfo.FixedFlag & FixedFlag.FromColFixed) == 0)
            {
                range.FromCol = _addressInfo.FromCol + colOffset;
            }
            else
            {
                range.FromCol = _addressInfo.FromCol;
            }

            if ((_addressInfo.FixedFlag & FixedFlag.ToColFixed) == 0)
            {
                range.ToCol = _addressInfo.ToCol + colOffset;
            }
            else
            {
                range.ToCol = _addressInfo.ToCol;
            }
            exp._result = new AddressCompileResult(range, DataType.ExcelCellAddress, range);
            return exp;
        }
    }
}
