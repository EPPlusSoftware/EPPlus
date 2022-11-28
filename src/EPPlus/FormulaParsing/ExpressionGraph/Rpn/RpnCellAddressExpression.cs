using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Diagnostics;
using Operators = OfficeOpenXml.FormulaParsing.Excel.Operators.Operators;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    [DebuggerDisplay("CellAddressExpression: {ExpressionString}")]
    internal class RpnCellAddressExpression : RpnExpression
    {
        protected FormulaRangeAddress _addressInfo;
        protected bool _negate =false;
        internal RpnCellAddressExpression(FormulaRangeAddress addressInfo, bool negate) : base(addressInfo._context)
        {
            _addressInfo = addressInfo;
            _negate = negate;
        }
        public RpnCellAddressExpression(string address, ParsingContext ctx, short externalReferenceIx, short worksheetIx) : base(ctx)
        {
            _addressInfo = new FormulaRangeAddress() { ExternalReferenceIx= externalReferenceIx, WorksheetIx = worksheetIx < 0 ? ctx.CurrentCell.WorksheetIx : worksheetIx };
            ExcelCellBase.GetRowColFromAddress(address, out int row, out int col, out bool fixedRow, out bool fixedCol);
            _addressInfo.FromRow = _addressInfo.ToRow = row;
            _addressInfo.FromCol = _addressInfo.ToCol = col;
            _addressInfo.FixedFlag = fixedRow ? FixedFlag.FromRowFixed | FixedFlag.ToRowFixed : 0;
            _addressInfo.FixedFlag |= fixedCol ? FixedFlag.FromColFixed | FixedFlag.ToColFixed : 0;
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
            return new RpnCellAddressExpression(new FormulaRangeAddress(Context) { FromRow = row, ToRow = row, FromCol = col, ToCol = col, ExternalReferenceIx = _addressInfo.ExternalReferenceIx, WorksheetIx = _addressInfo.WorksheetIx }, _negate)
            {
                Status = Status,                
                Operator= Operator
            };
        }
        public override FormulaRangeAddress GetAddress() { return _addressInfo; }
    }
}
