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
        FormulaRangeAddress _addressInfo;
        bool _negate=false;
        public RpnCellAddressExpression(string address, ParsingContext ctx, short externalReferenceIx, short worksheetIx) : base(ctx)
        {
            _addressInfo = new FormulaRangeAddress() { ExternalReferenceIx= externalReferenceIx, WorksheetIx = worksheetIx };
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
                    var wsIx = _addressInfo.WorksheetIx < 0 ? Context.CurrentCell.WorksheetIx : _addressInfo.WorksheetIx;
                    _result = CompileResultFactory.Create(Context.Package.Workbook.Worksheets[wsIx].GetValueInner(_addressInfo.FromRow, _addressInfo.FromCol), 0, _addressInfo);
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
    }
}
