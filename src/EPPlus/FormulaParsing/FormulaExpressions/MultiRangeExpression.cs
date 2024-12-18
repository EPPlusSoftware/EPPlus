using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using System.Diagnostics;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    /// <summary>
    /// This class represents a range that contains more than one address, where the ranges are separated by a comma.
    /// </summary>
    [DebuggerDisplay("MultiRangeExpression: {_addressInfo.Address}")]
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
