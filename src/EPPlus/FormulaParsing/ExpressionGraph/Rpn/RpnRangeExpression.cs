using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Packaging.Ionic;
using OfficeOpenXml.Utils;
using System;
using System.Diagnostics;
using System.Threading;
using Operators = OfficeOpenXml.FormulaParsing.Excel.Operators.Operators;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    [DebuggerDisplay("RpnRangeExpression: {_addressInfo.Address}")]
    internal class RpnRangeExpression : RpnExpression
    {
        protected FormulaRangeAddress _addressInfo;
        protected bool _negate =false;
        internal RpnRangeExpression(CompileResult result, ParsingContext ctx, bool negate) : base(ctx)
        {
            _cachedCompileResult = result;
            _addressInfo = result.Address;
            _negate = negate;
        }
        internal RpnRangeExpression(FormulaRangeAddress address, bool negate) : base(address._context)
        {
            _addressInfo = address;
            _negate = negate;
        }
        public RpnRangeExpression(string address, ParsingContext ctx, short externalReferenceIx, int worksheetIx) : base(ctx)
        {
            _addressInfo = new FormulaRangeAddress(ctx) { ExternalReferenceIx= externalReferenceIx, WorksheetIx = worksheetIx < 0 ? ctx.CurrentCell.WorksheetIx : worksheetIx };
            ExcelCellBase.GetRowColFromAddress(address, out int fromRow, out int fromCol, out int toRow, out int toCol, out bool fixedFromRow, out bool fixedFromCol, out bool fixedToRow, out bool fixedToCol);
            _addressInfo.FromRow = fromRow==0 ? 1 : fromRow;
            _addressInfo.ToRow = toRow == 0 ? ExcelPackage.MaxRows : toRow;
            _addressInfo.FromCol = fromCol == 0 ? 1 : fromCol;
            _addressInfo.ToCol = toCol == 0 ? ExcelPackage.MaxColumns : toCol; 
            _addressInfo.FixedFlag = (fixedFromRow ? FixedFlag.FromRowFixed : 0) | (fixedToRow ? FixedFlag.ToRowFixed : 0) | (fixedFromCol ? FixedFlag.FromColFixed : 0) | (fixedToCol ? FixedFlag.ToColFixed : 0);
        }
        internal override ExpressionType ExpressionType => ExpressionType.CellAddress;
        public override CompileResult Compile()
        {
            if (_cachedCompileResult == null)
            {
                if(_addressInfo.ExternalReferenceIx < 1)
                {
                    if (_addressInfo.IsSingleCell)
                    {

                        var ws = Context.Package.Workbook.GetWorksheetByIndexInList(_addressInfo.WorksheetIx);
                        var v = ws.GetValueInner(_addressInfo.FromRow, _addressInfo.FromCol);
                        if (_negate)
                        {
                            v = DoNegate(v);
                        }
                        _cachedCompileResult = CompileResultFactory.Create(v, _addressInfo);
                        _cachedCompileResult.IsHiddenCell = ws.IsCellHidden(_addressInfo.FromRow, _addressInfo.FromCol);
                    }
                    else
                    {
                        _cachedCompileResult = new AddressCompileResult(new RangeInfo(_addressInfo, Context), DataType.ExcelRange, _addressInfo);
                    }
                }
                else
                {
                    var wb = Context.GetExternalWoorkbook(_addressInfo.ExternalReferenceIx);
                    IRangeInfo ri;
                    if(wb?.Package!=null)
                    {
                        var ws = wb?.Package.Workbook.GetWorksheetByIndexInList(_addressInfo.WorksheetIx);
                        ri=new RangeInfo(ws, _addressInfo.FromRow, _addressInfo.FromCol, _addressInfo.ToRow, _addressInfo.ToCol, Context);
                    }
                    else
                    {
                        ri = new EpplusExcelExternalRangeInfo(wb, _addressInfo, Context);
                    }

                    if (ri.IsMulti)
                    {
                        _cachedCompileResult = new AddressCompileResult(ri, DataType.ExcelRange, _addressInfo);
                    }
                    else
                    {
                        var v = ri.GetOffset(0, 0);
                        if (_negate)
                        {
                            v = DoNegate(v);
                        }

                        _cachedCompileResult = CompileResultFactory.Create(v, _addressInfo);
                    }
                }
            }
            return _cachedCompileResult;
        }

        private object DoNegate(object v)
        {
            if(ConvertUtil.IsNumericOrDate(v))
            {
                return ConvertUtil.GetValueDouble(v) * -1;
            }
            return v;
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
            var fa = new FormulaRangeAddress(Context)
            {
                ExternalReferenceIx = _addressInfo.ExternalReferenceIx,
                WorksheetIx = _addressInfo.WorksheetIx,
                FromRow = (_addressInfo.FixedFlag & FixedFlag.FromRowFixed) == FixedFlag.FromRowFixed ? _addressInfo.FromRow : _addressInfo.FromRow + row,
                ToRow = (_addressInfo.FixedFlag & FixedFlag.ToRowFixed) == FixedFlag.ToRowFixed ? _addressInfo.ToRow : _addressInfo.ToRow + row,
                FromCol = (_addressInfo.FixedFlag & FixedFlag.FromColFixed) == FixedFlag.FromColFixed ? _addressInfo.FromCol : _addressInfo.FromCol + col,
                ToCol = (_addressInfo.FixedFlag & FixedFlag.ToColFixed) == FixedFlag.ToColFixed ? _addressInfo.ToCol : _addressInfo.ToCol + col,
            };
            return new RpnRangeExpression(fa, _negate)
            {
                Status = Status,                
                Operator= Operator
            };
        }
        public override FormulaRangeAddress GetAddress() { return _addressInfo.Clone(); }
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
