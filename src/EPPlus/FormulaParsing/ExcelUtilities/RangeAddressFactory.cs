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
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    internal class RangeAddressFactory
    {
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly AddressTranslator _addressTranslator;
        private readonly IndexToAddressTranslator _indexToAddressTranslator;

        internal RangeAddressFactory(ExcelDataProvider excelDataProvider)
            : this(excelDataProvider, new AddressTranslator(excelDataProvider), new IndexToAddressTranslator(excelDataProvider, ExcelReferenceType.RelativeRowAndColumn))
        {
           
            
        }

        internal RangeAddressFactory(ExcelDataProvider excelDataProvider, AddressTranslator addressTranslator, IndexToAddressTranslator indexToAddressTranslator)
        {
            Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
            Require.That(addressTranslator).Named("addressTranslator").IsNotNull();
            Require.That(indexToAddressTranslator).Named("indexToAddressTranslator").IsNotNull();
            _excelDataProvider = excelDataProvider;
            _addressTranslator = addressTranslator;
            _indexToAddressTranslator = indexToAddressTranslator;
        }

        public RangeAddress Create(int col, int row)
        {
            return Create(string.Empty, col, row);
        }

        public RangeAddress Create(string worksheetName, int col, int row)
        {
            return new RangeAddress()
            {
                Address = _indexToAddressTranslator.ToAddress(col, row),
                Worksheet = worksheetName,
                FromCol = col,
                ToCol = col,
                FromRow = row,
                ToRow = row
            };
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="worksheetName">will be used if no worksheet name is specified in <paramref name="address"/></param>
        /// <param name="address">address of a range</param>
        /// <returns></returns>
        public RangeAddress Create(string worksheetName, string address)
        {
            Require.That(address).Named("range").IsNotNullOrEmpty();
            //var addressInfo = ExcelAddressInfo.Parse(address);
            var adr = new ExcelAddressBase(address);  
            var sheet = string.IsNullOrEmpty(adr.WorkSheetName) ? worksheetName : adr.WorkSheetName;
            var dim = _excelDataProvider.GetDimensionEnd(sheet);
            var rangeAddress = new RangeAddress()
            {
                Address = adr.Address,
                Worksheet = sheet,
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

        public RangeAddress Create(string range)
        {
            Require.That(range).Named("range").IsNotNullOrEmpty();
            //var addressInfo = ExcelAddressInfo.Parse(range);
            var adr = new ExcelAddressBase(range);
            if (adr.Table != null)
            {
                var a = _excelDataProvider.GetRange(adr.WorkSheetName, range).Address;
                //Convert the Table-style Address to an A1C1 address
                adr = new ExcelAddressBase(a._fromRow, a._fromCol, a._toRow, a._toCol);
                adr._ws = a._ws;                
            }
            var rangeAddress = new RangeAddress()
            {
                Address = adr.Address,
                Worksheet = adr.WorkSheetName ?? "",
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

        private void HandleSingleCellAddress(RangeAddress rangeAddress, ExcelAddressInfo addressInfo)
        {
            int col, row;
            _addressTranslator.ToColAndRow(addressInfo.StartCell, out col, out row);
            rangeAddress.FromCol = col;
            rangeAddress.ToCol = col;
            rangeAddress.FromRow = row;
            rangeAddress.ToRow = row;
        }

        private void HandleMultipleCellAddress(RangeAddress rangeAddress, ExcelAddressInfo addressInfo)
        {
            int fromCol, fromRow;
            _addressTranslator.ToColAndRow(addressInfo.StartCell, out fromCol, out fromRow);
            int toCol, toRow;
            _addressTranslator.ToColAndRow(addressInfo.EndCell, out toCol, out toRow, AddressTranslator.RangeCalculationBehaviour.LastPart);
            rangeAddress.FromCol = fromCol;
            rangeAddress.ToCol = toCol;
            rangeAddress.FromRow = fromRow;
            rangeAddress.ToRow = toRow;
        }
    }
}
