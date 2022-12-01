using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Packaging.Ionic;
using System;
using System.Diagnostics;
using Operators = OfficeOpenXml.FormulaParsing.Excel.Operators.Operators;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    [DebuggerDisplay("CellAddressExpression: {_addressInfo.Address}")]
    internal class RpnRangeExpression : RpnExpression
    {
        protected FormulaRangeAddress _addressInfo;
        protected bool _negate =false;
        internal RpnRangeExpression(FormulaRangeAddress addressInfo, bool negate) : base(addressInfo._context)
        {
            _addressInfo = addressInfo;
            _negate = negate;
        }
        public RpnRangeExpression(string address, ParsingContext ctx, short externalReferenceIx, short worksheetIx) : base(ctx)
        {
            _addressInfo = new FormulaRangeAddress() { ExternalReferenceIx= externalReferenceIx, WorksheetIx = worksheetIx < 0 ? ctx.CurrentCell.WorksheetIx : worksheetIx };
            ExcelCellBase.GetRowColFromAddress(address, out int fromRow, out int fromCol, out int toRow, out int toCol, out bool fixedFromRow, out bool fixedFromCol, out bool fixedToRow, out bool fixedToCol);
            _addressInfo.FromRow = fromRow;
            _addressInfo.ToRow = toRow;
            _addressInfo.FromCol = fromCol;
            _addressInfo.ToCol = toCol;
            _addressInfo.FixedFlag = (fixedFromRow ? FixedFlag.FromRowFixed : 0) | (fixedToRow ? FixedFlag.ToRowFixed : 0) | (fixedFromCol ? FixedFlag.FromColFixed : 0) | (fixedToCol ? FixedFlag.ToColFixed : 0);
        }
        internal override ExpressionType ExpressionType => ExpressionType.CellAddress;
        CompileResult _result;
        public override CompileResult Compile()
        {
            if (_result == null)
            {
                if(_addressInfo.ExternalReferenceIx < 1)
                {
                    _result = CompileResultFactory.Create(Context.Package.Workbook.Worksheets[_addressInfo.WorksheetIx].GetValueInner(_addressInfo.FromRow, _addressInfo.FromCol), 0, _addressInfo);
                }
            }
            return _result;
        }
        public override void Negate()
        {
            _negate = !_negate;
        }
        internal override RpnExpressionStatus Status
        {
            get;
            set;
        } = RpnExpressionStatus.IsAddress;
        internal override RpnExpression CloneWithOffset(int row, int col)
        {
            row += _addressInfo.FromRow;
            col += _addressInfo.FromCol;
            return new RpnRangeExpression(new FormulaRangeAddress(Context) { FromRow = row, ToRow = row, FromCol = col, ToCol = col, ExternalReferenceIx = _addressInfo.ExternalReferenceIx, WorksheetIx = _addressInfo.WorksheetIx }, _negate)
            {
                Status = Status,                
                Operator= Operator
            };
        }
        public override FormulaRangeAddress GetAddress() { return _addressInfo; }
        internal override void MergeAddress(string address)
        {
            ExcelCellBase.GetRowColFromAddress(address, out int fromRow, out int fromCol, out int toRow, out int toCol, out bool fixedFromRow, out bool fixedFromCol, out bool fixedToRow, out bool fixedToCol);
            if (_addressInfo.FromRow > fromRow)
            {
                _addressInfo.FromRow = fromRow;
                SetFixedFlag(fixedFromRow, FixedFlag.FromRowFixed);
            }
            if (_addressInfo.ToRow < toRow)
            {
                _addressInfo.ToRow = toRow;
                SetFixedFlag(fixedToRow, FixedFlag.ToRowFixed);
            }
            if (_addressInfo.FromCol > fromCol)
            {
                _addressInfo.FromCol = fromCol;
                SetFixedFlag(fixedFromCol, FixedFlag.FromColFixed);
            }
            if (_addressInfo.ToCol < toCol)
            {
                _addressInfo.ToCol = toCol;
                SetFixedFlag(fixedToCol, FixedFlag.ToColFixed);
            }
        }

        private void SetFixedFlag(bool setFlag, FixedFlag flag)
        {
            if (setFlag)
            {
                _addressInfo.FixedFlag &= ~flag;
            }
            else
            {
                _addressInfo.FixedFlag |= flag;
            }
        }
    }
}
