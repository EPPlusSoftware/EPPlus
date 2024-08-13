using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    [DebuggerDisplay("RpnRangeExpression: {_addressInfo.Address}")]
    internal class RangeExpression : Expression
    {
        protected FormulaRangeAddress _addressInfo;
        internal RangeExpression(CompileResult result, ParsingContext ctx) : base(ctx)
        {
            _cachedCompileResult = result;
            _addressInfo = result.Address;
        }
        internal RangeExpression(FormulaRangeAddress address) : base(address._context)
        {
            _addressInfo = address;
        }
        internal RangeExpression(string address, ParsingContext ctx, short externalReferenceIx, int worksheetIx) : base(ctx)
        {
            Init(new FormulaRangeAddress(ctx, address), ctx, externalReferenceIx, worksheetIx);
        }
        internal RangeExpression(ExcelAddressBase address, ParsingContext ctx, short externalReferenceIx, int worksheetIx) : base(ctx)
        {
            Init(address.AsFormulaRangeAddress(ctx), ctx, externalReferenceIx, worksheetIx);
        }

        private void Init(FormulaRangeAddress address, ParsingContext ctx, short externalReferenceIx, int worksheetIx)
        {
            _addressInfo = address;
            _addressInfo.ExternalReferenceIx = externalReferenceIx;
            _addressInfo.WorksheetIx = (worksheetIx == int.MinValue ? ctx.CurrentCell.WorksheetIx : worksheetIx);
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
                        if (_addressInfo.WorksheetIx == -1)
                        {
                            _cachedCompileResult = CompileResult.GetErrorResult(eErrorType.Ref);
                        }
                        else
                        {
                            var ws = Context.Package.Workbook.GetWorksheetByIndexInList(_addressInfo.WorksheetIx);
                            var v = ws.GetValue(_addressInfo.FromRow, _addressInfo.FromCol); //Use GetValue to get richtext values.
                            _cachedCompileResult = CompileResultFactory.Create(v, _addressInfo);
                            _cachedCompileResult.IsHiddenCell = ws.IsRowHidden(_addressInfo.FromRow);
                        }
                    }
                    else
                    {
                        _cachedCompileResult = new AddressCompileResult(new RangeInfo(_addressInfo), DataType.ExcelRange, _addressInfo);
                    }
                }
                else
                {
                    var ri = _addressInfo.GetAsRangeInfo();
                    if (ri.GetNCells() > 1)
                    {
                        _cachedCompileResult = new AddressCompileResult(ri, DataType.ExcelRange, _addressInfo);
                    }
                    else
                    {
                        var v = ri.GetOffset(0, 0);
                        _cachedCompileResult = CompileResultFactory.Create(v, _addressInfo);
                    }
                }
            }
            return _cachedCompileResult;
        }

        public override Expression Negate()
        {
            if (_cachedCompileResult == null)
            {
                Compile();
            }
            return new RangeExpression(_cachedCompileResult.Negate(), Context);
        }
        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.IsAddress;
        internal override Expression CloneWithOffset(int row, int col)
        {
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
                Operator= Operator
            };
        }
        public override FormulaRangeAddress[] GetAddress() 
        {
            return [_addressInfo.Clone()];
        }
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
    internal class MultiRangeExpression : Expression
    {
        protected ExcelAddressBase _addressInfo;
        internal MultiRangeExpression(ExcelAddressBase address, ParsingContext ctx) : base(ctx)
        {
            _addressInfo = address;
        }
        internal override ExpressionType ExpressionType => ExpressionType.MultiAddress;
        public override CompileResult Compile()
        {
            if (_cachedCompileResult == null)
            {
                if (_addressInfo.ExternalReferenceIndex < 1)
                {
                    var ws = string.IsNullOrEmpty(_addressInfo.WorkSheetName) ? Context.CurrentWorksheet : Context.Package.Workbook.Worksheets[_addressInfo.WorkSheetName];
                    if (_addressInfo.IsSingleCell && _addressInfo.Addresses.Count==1)
                    {
                        if (string.IsNullOrEmpty(_addressInfo.WorkSheetName)==false && Context.GetWorksheetIndex(_addressInfo.WorkSheetName)<0)
                        {
                            _cachedCompileResult = CompileResult.GetErrorResult(eErrorType.Ref);
                        }
                        else
                        {
                            var v = ws.GetValue(_addressInfo._fromRow, _addressInfo._fromCol); //Use GetValue to get richtext values.
                            _cachedCompileResult = CompileResultFactory.Create(v, _addressInfo.AsFormulaRangeAddress(Context));
                            _cachedCompileResult.IsHiddenCell = ws.IsRowHidden(_addressInfo._fromRow);
                        }
                    }
                    else
                    {
                        _cachedCompileResult = new AddressCompileResult(new RangeInfo(ws, _addressInfo,Context), DataType.ExcelRange, _addressInfo.AsFormulaRangeAddress(Context));
                    }
                }
                else
                {
                    var fa = _addressInfo.AsFormulaRangeAddress(Context);
                    var ri = fa.GetAsRangeInfo();
                    if (ri.GetNCells() > 1)
                    {
                        _cachedCompileResult = new AddressCompileResult(ri, DataType.ExcelRange, fa);
                    }
                    else
                    {
                        var v = ri.GetOffset(0, 0);
                        _cachedCompileResult = CompileResultFactory.Create(v, fa);
                    }
                }
            }
            return _cachedCompileResult;
        }

        public override Expression Negate()
        {
            if (_cachedCompileResult == null)
            {
                Compile();
            }
            return new RangeExpression(_cachedCompileResult.Negate(), Context);
        }
        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.IsAddress;
        internal override Expression CloneWithOffset(int row, int col)
        {
            var ai = _addressInfo.CloneWithOffset(row, col);
            return new MultiRangeExpression(ai, Context)
            {
                Status = Status,
                Operator = Operator
            };
        }
        public override FormulaRangeAddress[] GetAddress()
        {
            var addresses = _addressInfo.GetAllAddresses();
            return addresses.Select(x=>x.AsFormulaRangeAddress(Context)).ToArray();
        }
        internal override void MergeAddress(string address)
        {
            int endIx=-1;
            var wb = ExcelAddress.GetWorkbookPart(address);
            var ws = ExcelAddress.GetWorksheetPart(address, null, ref endIx);
            if(endIx>0)
            {
                address = address.Substring(endIx);
            }
            ExcelCellBase.GetRowColFromAddress(address, out int fromRow, out int fromCol, out int toRow, out int toCol, out bool fixedFromRow, out bool fixedFromCol, out bool fixedToRow, out bool fixedToCol);            

            foreach (var sa in _addressInfo.GetAllAddresses())
            {
                if(string.IsNullOrEmpty(ws)==false && string.IsNullOrEmpty(sa.WorkSheetName)==false && !ws.Equals(sa.WorkSheetName,System.StringComparison.CurrentCultureIgnoreCase))
                {
                    _addressInfo=null;
                    return;
                }
                if (fromRow > sa._fromRow)
                {
                    fromRow = sa._fromRow;
                }
                if (toRow < sa._toRow)
                {
                    toRow = sa._toRow;
                }
                if(fromCol > sa._fromCol)
                {
                    fromRow = sa._fromCol;
                }
                if (toCol < sa._toCol)
                {
                    toCol = sa._toCol;
                }
                _addressInfo = new ExcelAddressBase(fromRow, fromCol, toRow, toCol);
            }
        }
    }
}
