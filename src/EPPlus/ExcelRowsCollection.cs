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
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml
{
    public class ExcelRowsCollection : ExcelRangeRow
    {
        ExcelWorksheet _worksheet;
        internal ExcelRowsCollection(ExcelWorksheet worksheet) : base(worksheet, 1, ExcelPackage.MaxRows)
        {
            _worksheet = worksheet;
        }
        public ExcelRangeRow this[int row]
        {
            get
            {
                return new ExcelRangeRow(_worksheet, row, row);
            }
        }
        public ExcelRangeRow this[int fromRow, int toRow]
        {
            get
            {            
                return new ExcelRangeRow(_worksheet, fromRow, toRow);
            }
        }        
    }
}