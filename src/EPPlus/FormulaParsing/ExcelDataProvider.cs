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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// This class should be implemented to be able to deliver excel data
    /// to the formula parser.
    /// </summary>
    public abstract class ExcelDataProvider : IDisposable
    {
        /// <summary>
        /// A range of cells in a worksheet.
        /// </summary>
        public interface IRangeInfo : IEnumerator<ICellInfo>, IEnumerable<ICellInfo>
        {
            /// <summary>
            /// If the range is empty
            /// </summary>
            bool IsEmpty { get; }
            /// <summary>
            /// If the contains more than one cell  with a value.
            /// </summary>
            bool IsMulti { get; }
            /// <summary>
            /// Get number of cells
            /// </summary>
            /// <returns>Number of cells</returns>
            int GetNCells();
            /// <summary>
            /// The address.
            /// </summary>
            ExcelAddressBase Address { get; }
            /// <summary>
            /// Get the value from a cell
            /// </summary>
            /// <param name="row">The Row</param>
            /// <param name="col">The Column</param>
            /// <returns></returns>
            object GetValue(int row, int col);
            /// <summary>
            /// Gets
            /// </summary>
            /// <param name="rowOffset"></param>
            /// <param name="colOffset"></param>
            /// <returns></returns>
            object GetOffset(int rowOffset, int colOffset);
            /// <summary>
            /// The worksheet 
            /// </summary>
            ExcelWorksheet Worksheet { get; }
        }
        /// <summary>
        /// Information and help methods about a cell
        /// </summary>
        public interface ICellInfo
        {
            string Address { get; }

            string WorksheetName { get; }
            int Row { get; }
            int Column { get; }

            ulong Id { get; }
            string Formula { get;  }
            object Value { get; }
            double ValueDouble { get; }
            double ValueDoubleLogical { get; }
            bool IsHiddenRow { get; }
            bool IsExcelError { get; }
            IList<Token> Tokens { get; }   
        }
        public interface INameInfo
        {
            ulong Id { get; set; }
            string Worksheet {get; set;}
            string Name { get; set; }
            string Formula { get; set; }
            IList<Token> Tokens { get; }
            object Value { get; set; }
        }

        /// <summary>
        /// Returns the names of the worksheets in a workbook
        /// </summary>
        /// <returns></returns>
        public abstract IEnumerable<string> GetWorksheets();
        /// <summary>
        /// Returns the names of all worksheet names
        /// </summary>
        /// <returns></returns>
        public abstract ExcelNamedRangeCollection GetWorksheetNames(string worksheet);
        /// <summary>
        /// Returns the names of all worksheet names
        /// </summary>
        /// <returns></returns>
        public abstract bool IsExternalName(string name);

        public abstract ExcelTable GetExcelTable(string name);
        /// <summary>
        /// Returns the number of a worksheet in the workbook
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <returns>The number within the workbook</returns>
        public abstract int GetWorksheetIndex(string worksheetName);

        /// <summary>
        /// Returns all defined names in a workbook
        /// </summary>
        /// <returns></returns>
        public abstract ExcelNamedRangeCollection GetWorkbookNameValues();
        /// <summary>
        /// Returns values from the required range.
        /// </summary>
        /// <param name="worksheetName">The name of the worksheet</param>
        /// <param name="row">Row</param>
        /// <param name="column">Column</param>
        /// <param name="address">The reference address</param>
        /// <returns></returns>
        public abstract IRangeInfo GetRange(string worksheetName, int row, int column, string address);
        /// <summary>
        /// Returns values from the required range.
        /// </summary>
        /// <param name="worksheetName">The name of the worksheet</param>
        /// <param name="address">The reference address</param>
        /// <returns></returns>
        public abstract IRangeInfo GetRange(string worksheetName, string address);
        public abstract INameInfo GetName(string worksheet, string name);

        public abstract IEnumerable<object> GetRangeValues(string address);

        public abstract string GetRangeFormula(string worksheetName, int row, int column);
        public abstract List<Token> GetRangeFormulaTokens(string worksheetName, int row, int column);
        public abstract bool IsRowHidden(string worksheetName, int row);
        ///// <summary>
        ///// Returns a single cell value
        ///// </summary>
        ///// <param name="address"></param>
        ///// <returns></returns>
        //public abstract object GetCellValue(int sheetID, string address);

        /// <summary>
        /// Returns a single cell value
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public abstract object GetCellValue(string sheetName, int row, int col);

        /// <summary>
        /// Creates a cell id, representing the full address of a cell.
        /// </summary>
        /// <param name="sheetName">Name of the worksheet</param>
        /// <param name="row">Row ix</param>
        /// <param name="col">Column Index</param>
        /// <returns>An <see cref="ulong"/> representing the addrss</returns>
        public abstract ulong GetCellId(string sheetName, int row, int col);

        ///// <summary>
        ///// Sets the value on the cell
        ///// </summary>
        ///// <param name="address"></param>
        ///// <param name="value"></param>
        //public abstract void SetCellValue(string address, object value);

        /// <summary>
        /// Returns the address of the lowest rightmost cell on the worksheet.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public abstract ExcelCellAddress GetDimensionEnd(string worksheet);

        /// <summary>
        /// Use this method to free unmanaged resources.
        /// </summary>
        public abstract void Dispose();

        /// <summary>
        /// Max number of columns in a worksheet that the Excel data provider can handle.
        /// </summary>
        public abstract int ExcelMaxColumns { get; }

        /// <summary>
        /// Max number of rows in a worksheet that the Excel data provider can handle
        /// </summary>
        public abstract int ExcelMaxRows { get; }

        public abstract object GetRangeValue(string worksheetName, int row, int column);
        public abstract string GetFormat(object value, string format);

        public abstract void Reset();
        public abstract IRangeInfo GetRange(string worksheet, int fromRow, int fromCol, int toRow, int toCol);
    }
}
