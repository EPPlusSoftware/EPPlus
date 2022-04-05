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
    /// <summary>
    /// A collection of columns in a worksheet
    /// </summary>
    public class ExcelColumnCollection : ExcelRangeColumn
    {
        ExcelWorksheet _worksheet;
        internal ExcelColumnCollection(ExcelWorksheet worksheet) : base(worksheet, 1, ExcelPackage.MaxColumns)
        {
            _worksheet = worksheet;            
            if(worksheet.Dimension!=null)
            {
                _fromCol = worksheet.Dimension._fromCol;
                _toCol = worksheet.Dimension._toCol;
            }
        }
        /// <summary>
        /// Indexer referenced by column index
        /// </summary>
        /// <param name="column">The column index</param>
        /// <returns>The column</returns>
        public ExcelRangeColumn this[int column]
        {
            get
            {
                return new ExcelRangeColumn(_worksheet, column, column);
            }
        }
        /// <summary>
        /// Indexer referenced by from and to column index
        /// </summary>
        /// <param name="fromColumn">Column from index</param>
        /// <param name="toColumn">Column to index</param>
        /// <returns></returns>
        public ExcelRangeColumn this[int fromColumn, int toColumn]
        {
            get
            {            
                return new ExcelRangeColumn(_worksheet, fromColumn, toColumn);
            }
        }        
    }
}