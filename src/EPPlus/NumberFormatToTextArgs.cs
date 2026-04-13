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
using OfficeOpenXml.Style.Interfaces;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;

namespace OfficeOpenXml
{
    /// <summary>
    /// Numberformat settings used in the <see cref="ExcelWorkbook.NumberFormatToTextHandler"/>
    /// </summary>
    public class NumberFormatToTextArgs
    {
        internal int _styleId;

        ExcelNumberFormatWithoutId fallbackNumberFormat = null;

        /// <summary>
        /// If these args are provided from a formula
        /// </summary>
        public bool FromFormula = false;

        /// <summary>
        /// Constructor when numberformat is not built in
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="value"></param>
        /// <param name="numberFormat"></param>
        internal NumberFormatToTextArgs(ExcelWorksheet ws, int row, int column, object value, string numberFormat)
        {
            Worksheet = ws;
            Row = row;
            Column = column;
            Value = value;
            _styleId = -1;
            fallbackNumberFormat = new ExcelNumberFormatWithoutId(numberFormat);
        }

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
        public IExcelNumberFormat NumberFormat 
        { 
            get 
            {               
                if(fallbackNumberFormat != null)
                {
                    return fallbackNumberFormat;
                }
                else
                {
                    return ValueToTextHandler.GetNumberFormat(_styleId, Worksheet.Workbook.Styles);
                }
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
                if(fallbackNumberFormat != null)
                {
                    var ft = new ExcelFormatTranslator(NumberFormat.Format, -1);
                    bool isValidFormat = false;
                    var frmt = ValueToTextHandler.FormatValue(Value, false, ft, null, out isValidFormat);
                    return frmt;
                }
                return ValueToTextHandler.GetFormattedText(Value, Worksheet.Workbook, _styleId, false);
            }
        }
    }
}