/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/27/2024         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml
{
    /// <summary>
    /// Numberformat settings used in the <see cref="ExcelWorkbook.NumberFormatToTextHandler"/>
    /// </summary>
    public class NumberFormatToTextArgs
    {
        internal int _styleId;
        internal NumberFormatToTextArgs(ExcelWorksheet ws, int row, int column, object value, int styleId)
        {
            Worksheet = ws;
            Row = row;
            Column = column;
            Value = value;
            _styleId = styleId;            
        }
        /// <summary>
        /// The worksheet of the cell.
        /// </summary>
        public ExcelWorksheet Worksheet { get; }
        /// <summary>
        /// The Row of the cell.
        /// </summary>
        public int Row { get; }
        /// <summary>
        /// The column of the cell.
        /// </summary>
        public int Column { get;  }
        /// <summary>
        /// The number format settings for the cell
        /// </summary>
        public ExcelNumberFormatXml NumberFormat 
        { 
            get 
            {
                return ValueToTextHandler.GetNumberFormat(_styleId, Worksheet.Workbook.Styles);
            } 
        } 
        /// <summary>
        /// The value of the cell to be formatted
        /// </summary>
        public object Value { get; }
        /// <summary>
        /// The text formatted by EPPlus
        /// </summary>
        public string Text
        {
            get
            { 
                return ValueToTextHandler.GetFormattedText(Value, Worksheet.Workbook, _styleId, false);
            }
        }
    }
}