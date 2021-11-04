/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
namespace OfficeOpenXml.ExternalReferences
{
    /// <summary>
    /// Represents a cell value of an external worksheets cell.
    /// </summary>
    public class ExcelExternalCellValue
    {
        /// <summary>
        /// The address of the cell
        /// </summary>
        public string Address 
        { 
            get 
            { 
                return ExcelCellBase.GetAddress(Row, Column); 
            } 
        }
        /// <summary>
        /// The row of the cell
        /// </summary>
        public int Row { get; internal set; }
        /// <summary>
        /// The column of the cell
        /// </summary>
        public int Column { get; internal set; }
        /// <summary>
        /// The value of the cell
        /// </summary>
        public object Value { get; internal set; }
        /// <summary>
        /// A reference index to meta data for the cell
        /// </summary>
        public int MetaDataReference { get; internal set; }
    }
}
