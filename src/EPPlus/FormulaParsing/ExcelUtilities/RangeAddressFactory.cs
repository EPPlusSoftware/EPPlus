/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    internal class RangeAddressFactory
    {
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly AddressTranslator _addressTranslator;
        private readonly IndexToAddressTranslator _indexToAddressTranslator;
        private readonly ParsingContext _context;

        internal RangeAddressFactory(ExcelDataProvider excelDataProvider, ParsingContext context)
            : this(excelDataProvider, new AddressTranslator(excelDataProvider), new IndexToAddressTranslator(excelDataProvider, ExcelReferenceType.RelativeRowAndColumn), context)
        {
           
            
        }

        internal RangeAddressFactory(ExcelDataProvider excelDataProvider, AddressTranslator addressTranslator, IndexToAddressTranslator indexToAddressTranslator, ParsingContext context)
        {
            Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
            Require.That(addressTranslator).Named("addressTranslator").IsNotNull();
            Require.That(indexToAddressTranslator).Named("indexToAddressTranslator").IsNotNull();
            Require.That(context).Named("context").IsNotNull();
            _excelDataProvider = excelDataProvider;
            _addressTranslator = addressTranslator;
            _indexToAddressTranslator = indexToAddressTranslator;
            _context = context;
        }

        public FormulaRangeAddress Create(int col, int row)
        {
            return Create(int.MinValue, col, row);
        }

        public FormulaRangeAddress Create(int wsIx, int col, int row)
        {
            return new FormulaRangeAddress(_context)
            {
                WorksheetIx = wsIx,
                FromCol = col,
                ToCol = col,
                FromRow = row,
                ToRow = row
            };
        }

        /// <summary>
        /// Create
        /// </summary>
        /// <param name="wsIx">The worksheet index.</param>
        /// <param name="address">address of a range</param>
        /// <returns></returns>
        public FormulaRangeAddress Create(int wsIx, string address)
        {
            Require.That(address).Named("range").IsNotNullOrEmpty();
            var adr = new ExcelAddressBase(address);  
            var worksheetIx = string.IsNullOrEmpty(adr.WorkSheetName) ? wsIx : _context.GetWorksheetIndex(adr.WorkSheetName);            
            var dim = _excelDataProvider.GetDimensionEnd((int)wsIx);
            var rangeAddress = new FormulaRangeAddress(_context)
            {
                WorksheetIx = worksheetIx,
                FromRow = adr._fromRow,
                FromCol = adr._fromCol,
                ToRow = (dim != null && adr._toRow > dim.Row) ? dim.Row : adr._toRow,
                ToCol = adr._toCol
            };

            //if (addressInfo.IsMultipleCells)
            //{
            //    HandleMultipleCellAddress(rangeAddress, addressInfo);
            //}
            //else
            //{
            //    HandleSingleCellAddress(rangeAddress, addressInfo);
            //}
            return rangeAddress;
        }

        public FormulaRangeAddress Create(string range)
        {
            Require.That(range).Named("range").IsNotNullOrEmpty();
            //var addressInfo = ExcelAddressInfo.Parse(range);
            var adr = new ExcelAddressBase(range);
            if (adr.Table != null)
            {
                var a = _excelDataProvider.GetRange(adr.WorkSheetName, range).Address;
                //Convert the Table-style Address to an A1C1 address
                adr = new ExcelAddressBase(a.FromRow, a.FromCol, a.ToRow, a.ToCol);
                adr._ws = a.WorksheetName;                
            }
            int worksheetIx = int.MinValue;
            if (!string.IsNullOrEmpty(adr._ws))
            {
                if (_context.Package != null && _context.Package.Workbook.Worksheets[adr._ws] != null)
                {
                    worksheetIx = (short)_context.Package.Workbook.Worksheets[adr._ws].PositionId;
                }
                else
                {
                    worksheetIx = -1;
                }
            }
            var rangeAddress = new FormulaRangeAddress(_context)
            {
                WorksheetIx = worksheetIx,
                FromRow = adr._fromRow,
                FromCol = adr._fromCol,
                ToRow = adr._toRow,
                ToCol = adr._toCol
            };
           
            //if (addressInfo.IsMultipleCells)
            //{
            //    HandleMultipleCellAddress(rangeAddress, addressInfo);
            //}
            //else
            //{
            //    HandleSingleCellAddress(rangeAddress, addressInfo);
            //}
            return rangeAddress;
        }
        public FormulaCellAddress CreateCell(string cell)
        {
            Require.That(cell).Named("cell").IsNotNullOrEmpty();
            ExcelAddressBase.GetRowColFromAddress(cell, out int row, out int col);            
            var wsIx = _excelDataProvider.GetWorksheetIndex(ExcelAddress.GetWorksheetPart(cell, "")) - 1;
            return new FormulaCellAddress(wsIx, row, col);
        }
    }
}
