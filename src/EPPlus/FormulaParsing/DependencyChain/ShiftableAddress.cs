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
namespace OfficeOpenXml
{
    internal class ShiftableAddress
    {
        int _fromRow, _fromCol, _toRow, _toCol;
        bool _fixedFromRow, _fixedToRow, _fixedFromCol, _fixedToCol;
        string _leftPart;
        public ShiftableAddress(string address)
        {
            int endIx = 0;
            ExcelAddressBase.GetWorksheetPart(address, "", ref endIx);
            if(endIx > 0)
            {
                _leftPart= address.Substring(0, endIx) + "!";
            }
            else
            {
                _leftPart = "";
            }
            ExcelCellBase.GetRowColFromAddress(address, out _fromRow, out _fromCol, out _toRow, out _toCol, out _fixedFromRow, out _fixedFromCol, out _fixedToRow, out _fixedToCol);
        }
        internal string GetOffsetAddress(int rowOffset, int colOffset)
        {
            int fromRow = _fixedFromRow ? _fromRow : _fromRow + rowOffset;
            int fromCol = _fixedFromCol ? _fromCol : _fromCol + colOffset;
            int toRow = _fixedToRow ? _toRow : _toRow + rowOffset;
            int toCol = _fixedToCol ? _toCol : _toCol + colOffset;
            return _leftPart + ExcelCellBase.GetAddress(fromRow, fromCol, toRow, toCol, _fixedFromRow, _fixedFromCol, _fixedToRow, _fixedToCol);
        }
    }
}