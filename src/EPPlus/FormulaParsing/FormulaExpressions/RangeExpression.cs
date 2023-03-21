using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;
using System.Diagnostics;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    [DebuggerDisplay("RpnRangeExpression: {_addressInfo.Address}")]
    internal class RangeExpression : Expression
    {
        protected FormulaRangeAddress _addressInfo;
        protected int _negate;
        internal RangeExpression(CompileResult result, ParsingContext ctx) : base(ctx)
        {
            _cachedCompileResult = result;
            _addressInfo = result.Address;
            _negate = 0;
        }
        internal RangeExpression(FormulaRangeAddress address, int negate) : base(address._context)
        {
            _addressInfo = address;
            _negate = negate;
        }
        public RangeExpression(string address, ParsingContext ctx, short externalReferenceIx, int worksheetIx) : base(ctx)
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
                        var v = ws.GetValue(_addressInfo.FromRow, _addressInfo.FromCol); //Use GetValue to get richtext values.
                        _cachedCompileResult = GetNegatedValue(v, _addressInfo);                       
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
                        _cachedCompileResult = GetNegatedValue(v, _addressInfo);                        
                    }
                }
            }
            return _cachedCompileResult;
        }

        private CompileResult GetNegatedValue(object value, FormulaRangeAddress addressInfo)
        {
            if (_negate == 0)
            {
                return CompileResultFactory.Create(value, addressInfo);
            }
            else
            {
                var d = ConvertUtil.GetValueDouble(value, false, true);
                if (double.IsNaN(d))
                {
                    return CompileResultFactory.Create(ExcelErrorValue.Create(eErrorType.Value), addressInfo);
                }
                else
                {
                    return CompileResultFactory.Create(d * _negate, addressInfo);
                }
            }
        }
        public override void Negate()
        {
            if (_negate == 0)
            {
                _negate = -1;
            }
            else
            {
                _negate *= -1;
            }
        }
        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.IsAddress;
        internal override Expression CloneWithOffset(int row, int col)
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
            return new RangeExpression(fa, _negate)
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
                _addressInfo.FixedFlag |= flag;
            }
            else
            {
                _addressInfo.FixedFlag &= ~flag;
            }
        }
    }
}
