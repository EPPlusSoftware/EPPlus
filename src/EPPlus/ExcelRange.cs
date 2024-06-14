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
using System.Text;
using OfficeOpenXml.Style;
using System.Data;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Table;
namespace OfficeOpenXml
{
    /// <summary>
    /// A range of cells. 
    /// </summary>
    public class ExcelRange : ExcelRangeBase
    {
        #region "Constructors"
        internal ExcelRange(ExcelWorksheet sheet, string address)
            : base(sheet, address)
        {

        }
        internal ExcelRange(ExcelWorksheet sheet, int fromRow, int fromCol, int toRow, int toCol)
            : base(sheet)
        {
            _fromRow = fromRow;
            _fromCol = fromCol;
            _toRow = toRow;
            _toCol = toCol;
        }
        #endregion
        #region "Indexers"
        /// <summary>
        /// Access the range using an address
        /// </summary>
        /// <param name="Address">The address</param>
        /// <returns>A range object</returns>
        public ExcelRange this[string Address]
        {
            get
            {
                if (_worksheet.Names.ContainsKey(Address))
                {
                    if (_worksheet.Names[Address].IsName)
                    {
                        return null;
                    }
                    else
                    {
                        base.Address = _worksheet.Names[Address].Address;
                    }
                }
                else
                {
                    if(Address.IndexOfAny(new char[] { '\'', '[', '!' })>=0)
                    {
                        var a = new ExcelAddress(Address);
                        if(a.WorkSheetName!=null && a.WorkSheetName.Equals(_worksheet.Name, StringComparison.InvariantCultureIgnoreCase)==false)
                        {
                            throw new InvalidOperationException($"The worksheet address {Address} is not within the worksheet {_worksheet.Name}");
                        }
                    }
                    SetAddress(Address, _workbook, _worksheet.Name);
                    ChangeAddress();
                }
                if((_fromRow < 1 || _fromCol < 1) && Address.Equals("#REF!", StringComparison.InvariantCultureIgnoreCase)==false)
                {
                    throw (new InvalidOperationException("Address is not valid."));
                }
                _rtc = null;
                return this;
            }
        }

        private ExcelRange GetTableAddess(ExcelWorksheet _worksheet, string address)
        {
            int ixStart = address.IndexOf('[');
            if (ixStart == 0) //External Address
            {
                int ixEnd = address.IndexOf(']',ixStart+1);
                if (ixStart >= 0 & ixEnd >= 0)
                {
                    var external = address.Substring(ixStart + 1, ixEnd - 1);
                    //if (Worksheet.Workbook._externalReferences.Count < external)
                    //{
                    //foreach(var 
                    //}
                }
            }
            return null;
        }
        /// <summary>
        /// Access a single cell
        /// </summary>
        /// <param name="Row">The row</param>
        /// <param name="Col">The column</param>
        /// <returns>A range object</returns>
        public ExcelRange this[int Row, int Col]
        {
            get
            {
                ValidateRowCol(Row, Col);

                _fromCol = Col;
                _fromRow = Row;
                _toCol = Col;
                _toRow = Row;
                _rtc = null;
                // avoid address re-calculation
                //base.Address = GetAddress(_fromRow, _fromCol);
                _start = null;
                _end = null;
                _addresses = null;
                _address = GetAddress(_fromRow, _fromCol);
                ChangeAddress();
                return this;
            }
        }
        /// <summary>
        /// Access a range of cells
        /// </summary>
        /// <param name="FromRow">Start row</param>
        /// <param name="FromCol">Start column</param>
        /// <param name="ToRow">End Row</param>
        /// <param name="ToCol">End Column</param>
        /// <returns></returns>
        public ExcelRange this[int FromRow, int FromCol, int ToRow, int ToCol]
        {
            get
            {
                ValidateRowCol(FromRow, FromCol);
                ValidateRowCol(ToRow, ToCol);

                _fromCol = FromCol;
                _fromRow = FromRow;
                _toCol = ToCol;
                _toRow = ToRow;
                _rtc = null;
                // avoid address re-calculation
                //base.Address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
                _start = null;
                _end = null;
                _addresses = null;
                _address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
                ChangeAddress();
                return this;
            }
        }
        #endregion
        private static void ValidateRowCol(int Row, int Col)
        {
            if (Row < 1 || Row > ExcelPackage.MaxRows)
            {
                throw (new ArgumentException("Row out of range"));
            }
            if (Col < 1 || Col > ExcelPackage.MaxColumns)
            {
                throw (new ArgumentException("Column out of range"));
            }
        }
		
        /// <summary>
        /// Set the formula for the range.
        /// </summary>
        /// <param name="formula">The formula for the range.</param>
        /// <param name="asSharedFormula">If the formula should be created as a shared formula. If false each cell will be set individually. This can be useful if the formula returns a dynamic array result.</param>
        public void SetFormula(string formula, bool asSharedFormula = true)
		{
			if(asSharedFormula || IsName || formula == null || formula.Trim() == "")
            {
                Formula = formula;
            }
            else
            {
                for(int c=_fromCol; c<=_toCol; c++) 
                { 
                   for(int r=_fromRow; r<=_toRow; r++)
                    {
						Set_Formula(this, formula, r, c);
					}
				}
            }
		}
    }
}
