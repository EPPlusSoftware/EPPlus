using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static OfficeOpenXml.ExcelAddressBase;

namespace OfficeOpenXml
{
    internal class ExcelAddressCollideUtility
    {
        public ExcelAddressCollideUtility(ExcelAddressBase address)
        {
            _fromRow = address._fromRow;
            _toRow = address._toRow;
            _fromCol = address._fromCol;
            _toCol = address._toCol;
            _worksheetName = address.WorkSheetName;
        }

        public ExcelAddressCollideUtility(FormulaRangeAddress address, ParsingContext ctx)
        {
            _fromRow = address.FromRow;
            _toRow = address.ToRow;
            _fromCol = address.FromCol;
            _toCol = address.ToCol;
            _worksheetName = GetWsName(address, ctx);
        }

        private readonly int _fromRow, _toRow, _fromCol, _toCol;
        private readonly string _worksheetName;

        private static string GetWsName(FormulaRangeAddress address, ParsingContext ctx)
        {
            if (ctx.Package != null && ctx.Package.Workbook.Worksheets[address.WorksheetIx] != null)
            {
                return ctx.Package.Workbook.Worksheets[address.WorksheetIx].Name;
            }
            else
            {
                return address.WorksheetIx.ToString();
            }
        }

        internal eAddressCollition Collide(ExcelAddressBase address, bool ignoreWs = false)
        {
            if (ignoreWs == false && address.WorkSheetName != _worksheetName &&
                string.IsNullOrEmpty(address.WorkSheetName) == false &&
                string.IsNullOrEmpty(_worksheetName) == false)
            {
                return eAddressCollition.No;
            }

            return Collide(address._fromRow, address._fromCol, address._toRow, address._toCol);
        }

        internal eAddressCollition Collide(FormulaRangeAddress address, ParsingContext ctx, bool ignoreWs = false)
        {
            var ws = GetWsName(address, ctx);
            if (ignoreWs == false && ws != _worksheetName &&
                string.IsNullOrEmpty(ws) == false &&
                string.IsNullOrEmpty(_worksheetName) == false)
            {
                return eAddressCollition.No;
            }

            return Collide(address.FromRow, address.FromCol, address.ToRow, address.ToCol);
        }

        internal eAddressCollition Collide(int row, int col)
        {
            return Collide(row, col, row, col);
        }
        internal eAddressCollition Collide(int fromRow, int fromCol, int toRow, int toCol)
        {
            if (DoNotCollide(fromRow, fromCol, toRow, toCol))
            {
                return eAddressCollition.No;
            }
            else if (fromRow == _fromRow && fromCol == _fromCol &&
                    toRow == _toRow && toCol == _toCol)
            {
                return eAddressCollition.Equal;
            }
            else if (fromRow >= _fromRow && toRow <= _toRow &&
                     fromCol >= _fromCol && toCol <= _toCol)
            {
                return eAddressCollition.Inside;
            }
            else
                return eAddressCollition.Partly;
        }

        internal bool DoNotCollide(int fromRow, int fromCol, int toRow, int toCol)
        {
            return fromRow > _toRow || fromCol > _toCol
                   ||
                   _fromRow > toRow || _fromCol > toCol;
        }

        internal bool CollideFullRowOrColumn(ExcelAddressBase address)
        {
            return CollideFullRowOrColumn(address._fromRow, address._fromCol, address._toRow, address._toCol);
        }

        internal bool CollideFullRowOrColumn(FormulaRangeAddress address)
        {
            return CollideFullRowOrColumn(address.FromRow, address.FromCol, address.ToRow, address.ToCol);
        }
        internal bool CollideFullRowOrColumn(int fromRow, int fromCol, int toRow, int toCol)
        {
            return (CollideFullRow(fromRow, toRow) && CollideColumn(fromCol, toCol)) ||
                   (CollideFullColumn(fromCol, toCol) && CollideRow(fromRow, toRow));
        }
        private bool CollideColumn(int fromCol, int toCol)
        {
            return fromCol <= _toCol && toCol >= _fromCol;
        }

        internal bool CollideRow(int fromRow, int toRow)
        {
            return fromRow <= _toRow && toRow >= _fromRow;
        }
        internal bool CollideFullRow(int fromRow, int toRow)
        {
            return fromRow <= _fromRow && toRow >= _toRow;
        }
        internal bool CollideFullColumn(int fromCol, int toCol)
        {
            return fromCol <= _fromCol && toCol >= _toCol;
        }
    }
}
