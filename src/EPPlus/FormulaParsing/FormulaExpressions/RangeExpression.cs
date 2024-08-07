using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    [DebuggerDisplay("RpnRangeExpression: {_addressInfo.Address}")]
    internal class RangeExpression : Expression
    {
        protected FormulaRangeAddress[] _addressInfo;
        internal RangeExpression(CompileResult result, ParsingContext ctx) : base(ctx)
        {
            _cachedCompileResult = result;
            _addressInfo = [result.Address];
        }
        internal RangeExpression(FormulaRangeAddress address) : base(address._context)
        {
            _addressInfo = [address];
        }
        internal RangeExpression(FormulaRangeAddress[] address) : base(address.Length>0?address[0]._context : null)
        {
            _addressInfo = address;
        }
        internal RangeExpression(ExcelAddressBase address, ParsingContext ctx, short externalReferenceIx, int worksheetIx) : base(ctx)
        {
            if(address.Addresses == null || address.Addresses.Count==1)
            {
                //_addressInfo = new ;
            }
        }
        public RangeExpression(string address, ParsingContext ctx, short externalReferenceIx, int worksheetIx) : base(ctx)
        {
            //_addressInfo = new FormulaRangeAddress(ctx) { ExternalReferenceIx= externalReferenceIx, WorksheetIx = worksheetIx == int.MinValue ? ctx.CurrentCell.WorksheetIx : worksheetIx };
            var ab = new ExcelAddressBase(address);
            if(ab.Address==null)
            {
                _addressInfo = [ab.AsFormulaRangeAddress(ctx)];
            }
            else
            {
                _addressInfo = new FormulaRangeAddress[ab.Addresses.Count];
                var i = 0;
                foreach (var a in ab.Addresses)
                {
                    _addressInfo[i++]=a.AsFormulaRangeAddress(ctx);
                }
            }
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
                    var ri = _addressInfo[0].GetAsRangeInfo();
                    if (ri.GetNCells()>1)
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
            var ai = new FormulaRangeAddress[_addressInfo.Length];
            var i = 0;
            foreach(var fa in _addressInfo)
            {
                ai[i++] = new FormulaRangeAddress(Context)
                {
                    ExternalReferenceIx = fa.ExternalReferenceIx,
                    WorksheetIx = fa.WorksheetIx,
                    FromRow = (fa.FixedFlag & FixedFlag.FromRowFixed) == FixedFlag.FromRowFixed ? fa.FromRow : fa.FromRow + row,
                    ToRow = (fa.FixedFlag & FixedFlag.ToRowFixed) == FixedFlag.ToRowFixed ? fa.ToRow : fa.ToRow + row,
                    FromCol = (fa.FixedFlag & FixedFlag.FromColFixed) == FixedFlag.FromColFixed ? fa.FromCol : fa.FromCol + col,
                    ToCol = (fa.FixedFlag & FixedFlag.ToColFixed) == FixedFlag.ToColFixed ? fa.ToCol : fa.ToCol + col,
                };                
            }

            return new RangeExpression(ai)
            {
                Status = Status,                
                Operator= Operator
            };
        }
        public override Queue<FormulaRangeAddress> GetAddress() 
        {
            var q = new Queue<FormulaRangeAddress>();
            foreach (var a in _addressInfo)
            {
                q.Enqueue(a.Clone());
            }
            return q; 
        }
        internal override void MergeAddress(string address)
        {
            ExcelCellBase.GetRowColFromAddress(address, out int fromRow, out int fromCol, out int toRow, out int toCol, out bool fixedFromRow, out bool fixedFromCol, out bool fixedToRow, out bool fixedToCol);

            if (_addressInfo[0].FromRow > fromRow)
            {
                _addressInfo[0].FromRow = fromRow;
                SetFixedFlag(fixedFromRow, FixedFlag.FromRowFixed);
            }
            if (_addressInfo[0].ToRow < toRow)
            {
                _addressInfo[0].ToRow = toRow;
                SetFixedFlag(fixedToRow, FixedFlag.ToRowFixed);
            }
            if (_addressInfo[0].FromCol > fromCol)
            {
                _addressInfo[0].FromCol = fromCol;
                SetFixedFlag(fixedFromCol, FixedFlag.FromColFixed);
            }
            if (_addressInfo[0].ToCol < toCol)
            {
                _addressInfo[0].ToCol = toCol;
                SetFixedFlag(fixedToCol, FixedFlag.ToColFixed);
            }
        }

        private void SetFixedFlag(bool setFlag, FixedFlag flag)
        {
            if (setFlag)
            {
                _addressInfo[0].FixedFlag |= flag;
            }
            else
            {
                _addressInfo[0].FixedFlag &= ~flag;
            }
        }
    }
}
