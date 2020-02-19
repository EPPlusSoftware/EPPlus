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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Core;

namespace OfficeOpenXml
{
    /// <summary>
    /// Base class containing cell address manipulating methods.
    /// </summary>
    public abstract class ExcelCellBase
    {
        #region "Public Functions"
        /// <summary>
        /// Get the sheet, row and column from the CellID
        /// </summary>
        /// <param name="cellID"></param>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        static internal void SplitCellID(ulong cellID, out int sheet, out int row, out int col)
        {
            sheet = (int)(cellID % 0x8000);
            col = ((int)(cellID >> 15) & 0x3FF);
            row = ((int)(cellID >> 29));
        }
        /// <summary>
        /// Get the cellID for the cell. 
        /// </summary>
        /// <param name="SheetID"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        internal static ulong GetCellID(int SheetID, int row, int col)
        {
            return ((ulong)SheetID) + (((ulong)col) << 15) + (((ulong)row) << 29);
        }
        #endregion
        #region "Formula Functions"
        private delegate string dlgTransl(string part, int row, int col);
        #region R1C1 Functions"
        /// <summary>
        /// Translates a R1C1 to an absolut address/Formula
        /// </summary>
        /// <param name="value">Address</param>
        /// <param name="row">Current row</param>
        /// <param name="col">Current column</param>
        /// <returns>The RC address</returns>
        public static string TranslateFromR1C1(string value, int row, int col)
        {
            return R1C1Translator.FromR1C1Formula(value, row, col);
            //return Translate(value, ToAbs, row, col);
        }
        /// <summary>
        /// Translates a absolut address to R1C1 Format
        /// </summary>
        /// <param name="value">R1C1 Address</param>
        /// <param name="row">Current row</param>
        /// <param name="col">Current column</param>
        /// <returns>The absolut address/Formula</returns>
        public static string TranslateToR1C1(string value, int row, int col)
        {
            return R1C1Translator.ToR1C1Formula(value, row, col);
            //return Translate(value, ToR1C1, row, col);
        }
        //    //    return part;
        //    if (rStart != 0 && cStart != 0)
        #endregion
        #region "Address Functions"
        #region GetColumnLetter
        /// <summary>
        /// Returns the character representation of the numbered column
        /// </summary>
        /// <param name="iColumnNumber">The number of the column</param>
        /// <returns>The letter representing the column</returns>
        protected internal static string GetColumnLetter(int iColumnNumber)
        {
            return GetColumnLetter(iColumnNumber, false);
        }
        /// <summary>
        /// Returns the character representation of the numbered column
        /// </summary>
        /// <param name="iColumnNumber">The number of the column</param>
        /// <param name="fixedCol">True for fixed column</param>
        /// <returns>The letter representing the column</returns>
        protected internal static string GetColumnLetter(int iColumnNumber, bool fixedCol)
        {

            if (iColumnNumber < 1)
            {
                //throw new Exception("Column number is out of range");
                return "#REF!";
            }

            string sCol = "";
            do
            {
                sCol = ((char)('A' + ((iColumnNumber - 1) % 26))).ToString() + sCol;
                iColumnNumber = (iColumnNumber - ((iColumnNumber - 1) % 26)) / 26;
            }
            while (iColumnNumber > 0);
            return fixedCol ? "$" + sCol : sCol;
        }
        #endregion

        internal static bool GetRowColFromAddress(string CellAddress, out int FromRow, out int FromColumn, out int ToRow, out int ToColumn)
        {
            return GetRowColFromAddress(CellAddress, out FromRow, out FromColumn, out ToRow, out ToColumn, out _, out _, out _, out _);
        }
        /// <summary>
        /// Get the row/columns for a Cell-address
        /// </summary>
        /// <param name="CellAddress">The address</param>
        /// <param name="FromRow">Returns the to column</param>
        /// <param name="FromColumn">Returns the from column</param>
        /// <param name="ToRow">Returns the to row</param>
        /// <param name="ToColumn">Returns the from row</param>
        /// <param name="fixedFromRow">Is the from row fixed?</param>
        /// <param name="fixedFromColumn">Is the from column fixed?</param>
        /// <param name="fixedToRow">Is the to row fixed?</param>
        /// <param name="fixedToColumn">Is the to column fixed?</param>
        /// <returns></returns>
        internal static bool GetRowColFromAddress(string CellAddress, out int FromRow, out int FromColumn, out int ToRow, out int ToColumn, out bool fixedFromRow, out bool fixedFromColumn, out bool fixedToRow, out bool fixedToColumn)
        {
            bool ret;
            if (CellAddress.IndexOf('[') > 0) //External reference or reference to Table or Pivottable.
            {
                FromRow = -1;
                FromColumn = -1;
                ToRow = -1;
                ToColumn = -1;
                fixedFromRow = false;
                fixedFromColumn = false;
                fixedToRow = false;
                fixedToColumn = false;
                return false;
            }

            CellAddress = Utils.ConvertUtil._invariantTextInfo.ToUpper(CellAddress);
            //This one can be removed when the worksheet Select format is fixed
            if (CellAddress.IndexOf(' ') > 0)
            {
                CellAddress = CellAddress.Substring(0, CellAddress.IndexOf(' '));
            }

            if (CellAddress.IndexOf(':') < 0)
            {
                ret = GetRowColFromAddress(CellAddress, out FromRow, out FromColumn, out fixedFromRow, out fixedFromColumn);
                ToColumn = FromColumn;
                ToRow = FromRow;
                fixedToRow = fixedFromRow;
                fixedToColumn = fixedFromColumn;
            }
            else
            {
                string[] cells = CellAddress.Split(':');
                if(cells.Length>2)
                {
                    throw new InvalidOperationException($"Address is not valid {CellAddress}");
                }
                else
                {
                    if (IsCellAddress(cells[0]) != IsCellAddress(cells[1]))
                    {
                        throw new InvalidOperationException($"Address is not valid {CellAddress}");
                    }
                }
                ret = GetRowColFromAddress(cells[0], out FromRow, out FromColumn, out fixedFromRow, out fixedFromColumn);
                if (ret)
                    ret = GetRowColFromAddress(cells[1], out ToRow, out ToColumn, out fixedToRow, out fixedToColumn);
                else
                {
                    GetRowColFromAddress(cells[1], out ToRow, out ToColumn, out fixedToRow, out fixedToColumn);
                }

                if (FromColumn <= 0)
                    FromColumn = 1;
                if (FromRow <= 0)
                    FromRow = 1;
                if (ToColumn <= 0)
                    ToColumn = ExcelPackage.MaxColumns;
                if (ToRow <= 0)
                    ToRow = ExcelPackage.MaxRows;
            }
            return ret;
        }

        private static bool IsCellAddress(string cellAddress)
        {
            char start;
            if (cellAddress[0]=='$')
            {
                start = cellAddress.ToUpper()[1];
            }
            else
            {
                start = cellAddress.ToUpper()[0];
            }
            var end = cellAddress[cellAddress.Length-1];
            return (start >= 'A' && start <= 'Z' && end >= '0' && end <= '9');
        }

        /// <summary>
        /// Get the row/column for n Cell-address
        /// </summary>
        /// <param name="CellAddress">The address</param>
        /// <param name="Row">Returns Tthe row</param>
        /// <param name="Column">Returns the column</param>
        /// <returns>true if valid</returns>
        internal static bool GetRowColFromAddress(string CellAddress, out int Row, out int Column)
        {
            return GetRowCol(CellAddress, out Row, out Column, true);
        }
        internal static bool GetRowColFromAddress(string CellAddress, out int row, out int col, out bool fixedRow, out bool fixedCol)
        {
            return GetRowCol(CellAddress, out row, out col, true, out fixedRow, out fixedCol);
        }
        internal static bool IsAlpha(char c)
        {
            return c >= 'A' && c <= 'Z';
        }
        /// <summary>
        /// Get the row/column for a Cell-address
        /// </summary>
        /// <param name="address">the address</param>
        /// <param name="row">returns the row</param>
        /// <param name="col">returns the column</param>
        /// <param name="throwException">throw exception if invalid, otherwise returns false</param>
        /// <returns></returns>
        internal static bool GetRowCol(string address, out int row, out int col, bool throwException)
        {
            return GetRowCol(address, out row, out col, throwException, out bool fixedRow, out bool fixedCol);
        }
        const int numberOfCharacters = ('Z' - 'A') + 1;
        const int startChar = 'A'-1;
        const int startNum = '0';
        internal static bool GetRowCol(string address, out int row, out int col, bool throwException, out bool fixedRow, out bool fixedCol)
        {
            int start = 0;
            col = 0;
            row = 0;
            fixedRow = false;
            fixedCol = false;

            if (Utils.ConvertUtil._invariantCompareInfo.IsSuffix(address, "#REF!"))
            {
                row = 0;
                col = 0;
                return true;
            }

            var sheetNameSeparator = address.LastIndexOf('!');
            if (sheetNameSeparator > 0)
            {
                start = sheetNameSeparator + 1;
            }
            address = Utils.ConvertUtil._invariantTextInfo.ToUpper(address);
            for (int i = start; i < address.Length; i++)
            {
                char c = address[i];
                if (IsAlpha(c))
                {
                    col *= numberOfCharacters;
                    col += c - startChar;
                    if (col > ExcelPackage.MaxColumns || row > 0)
                    {
                        ThrowAddressException(address, out row, out col, throwException);
                        break;
                    }
                }
                else if (c >= '0' && c <= '9')
                {
                    row *= 10; //Number of numbers 0-9
                    row += c - startNum;
                    if (row > ExcelPackage.MaxRows)
                    {
                        ThrowAddressException(address, out row, out col, throwException);
                        break;  
                    }
                }
                else if (c == '$')
                {
                    if (IsAlpha(address[i + 1]))
                    {
                        fixedCol = true;
                    }
                    else
                    {
                        fixedRow = true;
                    }
                }
                else
                {
                    return ThrowAddressException(address, out row, out col, throwException);
                }
            }
            return row != 0 || col != 0;
        }

        private static bool ThrowAddressException(string address, out int row, out int col, bool throwException)
        {
            row = 0;
            col = 0;
            if (throwException)
            {
                throw (new ArgumentException(string.Format("Invalid Address format {0}", address)));
            }
            else
            {
                return false;
            }
        }

        private static int GetColumn(string sCol)
        {
            int col = 0;
            int len = sCol.Length - 1;
            for (int i = len; i >= 0; i--)
            {
                col += (((int)sCol[i]) - 64) * (int)(Math.Pow(26, len - i));
            }
            return col;
        }
        #region GetAddress
        /// <summary>
        /// Get the row number in text
        /// </summary>
        /// <param name="Row">The row</param>
        /// <param name="Absolute">If the row is absolute. Adds a $ before the address if true</param>
        /// <returns></returns>
        public static string GetAddressRow(int Row, bool Absolute = false)
        {
            if (Absolute)
                return $"${Row}";
            return $"{Row}";
        }
        /// <summary>
        /// Get the columnn address for the column
        /// </summary>
        /// <param name="Col">The column</param>
        /// <param name="Absolute">If the column is absolute. Adds a $ before the address if true</param>
        /// <returns></returns>
        public static string GetAddressCol(int Col, bool Absolute = false)
        {
            var colLetter = GetColumnLetter(Col);
            if (Absolute)
                return $"${colLetter}";
            return $"{colLetter}";
        }
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="Row">The number of the row</param>
        /// <param name="Column">The number of the column in the worksheet</param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int Row, int Column)
        {
            return GetAddress(Row, Column, false);
        }
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="Row">The number of the row</param>
        /// <param name="Column">The number of the column in the worksheet</param>
        /// <param name="AbsoluteRow">Absolute row</param>
        /// <param name="AbsoluteCol">Absolute column</param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int Row, bool AbsoluteRow, int Column, bool AbsoluteCol)
        {
            if (Row < 1 || Row > ExcelPackage.MaxRows || Column < 1 || Column > ExcelPackage.MaxColumns) return "#REF!";
            return (AbsoluteCol ? "$" : "") + GetColumnLetter(Column) + (AbsoluteRow ? "$" : "") + Row.ToString();
        }
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="Row">The number of the row</param>
        /// <param name="Column">The number of the column in the worksheet</param>
        /// <param name="Absolute">Get an absolute address ($A$1)</param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int Row, int Column, bool Absolute)
        {
            if (Row == 0 || Column == 0)
            {
                return "#REF!";
            }
            if (Absolute)
            {
                return ("$" + GetColumnLetter(Column) + "$" + Row.ToString());
            }
            else
            {
                return (GetColumnLetter(Column) + Row.ToString());
            }
        }
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="FromRow">From row number</param>
        /// <param name="FromColumn">From column number</param>
        /// <param name="ToRow">To row number</param>
        /// <param name="ToColumn">From column number</param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int FromRow, int FromColumn, int ToRow, int ToColumn)
        {
            return GetAddress(FromRow, FromColumn, ToRow, ToColumn, false);
        }
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="FromRow">From row number</param>
        /// <param name="FromColumn">From column number</param>
        /// <param name="ToRow">To row number</param>
        /// <param name="ToColumn">From column number</param>
        /// <param name="Absolute">if true address is absolute (like $A$1)</param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int FromRow, int FromColumn, int ToRow, int ToColumn, bool Absolute)
        {
            if (FromRow == ToRow && FromColumn == ToColumn)
            {
                return GetAddress(FromRow, FromColumn, Absolute);
            }
            else
            {
                if (FromRow == 1 && ToRow == ExcelPackage.MaxRows)
                {
                    var absChar = Absolute ? "$" : "";
                    return absChar + GetColumnLetter(FromColumn) + ":" + absChar + GetColumnLetter(ToColumn);
                }
                else if (FromColumn == 1 && ToColumn == ExcelPackage.MaxColumns)
                {
                    var absChar = Absolute ? "$" : "";
                    return absChar + FromRow.ToString() + ":" + absChar + ToRow.ToString();
                }
                else
                {
                    return GetAddress(FromRow, FromColumn, Absolute) + ":" + GetAddress(ToRow, ToColumn, Absolute);
                }
            }
        }
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="FromRow">From row number</param>
        /// <param name="FromColumn">From column number</param>
        /// <param name="ToRow">To row number</param>
        /// <param name="ToColumn">From column number</param>
        /// <param name="FixedFromColumn"></param>
        /// <param name="FixedFromRow"></param>
        /// <param name="FixedToColumn"></param>
        /// <param name="FixedToRow"></param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int FromRow, int FromColumn, int ToRow, int ToColumn, bool FixedFromRow, bool FixedFromColumn, bool FixedToRow, bool FixedToColumn)
        {
            if (FromRow == ToRow && FromColumn == ToColumn)
            {
                return GetAddress(FromRow, FixedFromRow, FromColumn, FixedFromColumn);
            }
            else
            {
                if (FromRow == 1 && ToRow == ExcelPackage.MaxRows)
                {
                    return GetColumnLetter(FromColumn, FixedFromColumn) + ":" + GetColumnLetter(ToColumn, FixedToColumn);
                }
                else if (FromColumn == 1 && ToColumn == ExcelPackage.MaxColumns)
                {
                    return (FixedFromRow ? "$" : "") + FromRow.ToString() + ":" + (FixedToRow ? "$" : "") + ToRow.ToString();
                }
                else
                {
                    return GetAddress(FromRow, FixedFromRow, FromColumn, FixedFromColumn) + ":" + GetAddress(ToRow, FixedToRow, ToColumn, FixedToColumn);
                }
            }
        }
        /// <summary>
        /// Get the full address including the worksheet name
        /// </summary>
        /// <param name="worksheetName">The name of the worksheet</param>
        /// <param name="address">The address</param>
        /// <returns>The full address</returns>
        public static string GetFullAddress(string worksheetName, string address)
        {
            return GetFullAddress(worksheetName, address, true);
        }
        /// <summary>
        /// Get the full address including the worksheet name
        /// </summary>
        /// <param name="workbook">The workbook, if other than current</param>   
        /// <param name="worksheetName">The name of the worksheet</param>
        /// <param name="address">The address</param>
        /// <returns>The full address</returns>
        public static string GetFullAddress(string workbook, string worksheetName, string address)
        {
            if (!string.IsNullOrEmpty(workbook))
                workbook = $"[{workbook}]";
            return workbook + GetFullAddress(worksheetName, address, true);
        }
        internal static string GetFullAddress(string worksheetName, string address, bool fullRowCol)
        {
            if (!string.IsNullOrEmpty(worksheetName)) worksheetName = worksheetName.Replace("'", "''");   //Makesure addresses handle single qoutes
            if (address.IndexOf("!") == -1 || address == "#REF!")
            {
                if (fullRowCol)
                {
                    string[] cells = address.Split(':');
                    if (cells.Length > 0)
                    {
                        address = string.IsNullOrEmpty(worksheetName) ? cells[0] : string.Format("'{0}'!{1}", worksheetName, cells[0]);
                        if (cells.Length > 1)
                        {
                            address += string.Format(":{0}", cells[1]);
                        }
                    }
                }
                else
                {
                    var a = new ExcelAddressBase(address);
                    if ((a._fromRow == 1 && a._toRow == ExcelPackage.MaxRows) || (a._fromCol == 1 && a._toCol == ExcelPackage.MaxColumns))
                    {
                        if (string.IsNullOrEmpty(worksheetName))
                        {
                            address = $"'{worksheetName}'!";
                        }
                        address += string.Format("{0}{1}:{2}{3}", ExcelAddress.GetColumnLetter(a._fromCol), a._fromRow, ExcelAddress.GetColumnLetter(a._toCol), a._toRow);
                    }
                    else
                    {
                        address = GetFullAddress(worksheetName, address, true);
                    }
                }
            }
            return address;
        }
        #endregion
        #region IsValidCellAddress
        /// <summary>
        /// Returns true if the cell address is valid
        /// </summary>
        /// <param name="address">The address to check</param>
        /// <returns>Return true if the address is valid</returns>
        public static bool IsValidAddress(string address)
        {
            if (address.LastIndexOf('!') > 0)
            {
                address = address.Substring(address.LastIndexOf('!') + 1);
            }
            if (string.IsNullOrEmpty(address.Trim())) return false;
            address = Utils.ConvertUtil._invariantTextInfo.ToUpper(address);
            var addrs = address.Split(',');
            foreach (var a in addrs)
            {
                string r1 = "", c1 = "", r2 = "", c2 = "";
                bool isSecond = false;
                for (int i = 0; i < a.Length; i++)
                {
                    if (IsCol(a[i]))
                    {
                        if (isSecond == false)
                        {
                            if (r1 != "") return false;
                            c1 += a[i];
                            if (c1.Length > 3) return false;
                        }
                        else
                        {
                            if (r2 != "") return false;
                            c2 += a[i];
                            if (c2.Length > 3) return false;
                        }
                    }
                    else if (IsRow(a[i]))
                    {
                        if (isSecond == false)
                        {
                            r1 += a[i];
                            if (r1.Length > 7) return false;
                        }
                        else
                        {
                            r2 += a[i];
                            if (r2.Length > 7) return false;
                        }
                    }
                    else if (a[i] == ':')
                    {
                        if (isSecond || i == a.Length - 1) return false;
                        isSecond = true;
                    }
                    else if (a[i] == '$')
                    {
                        if (i == a.Length - 1 || a[i + 1] == ':' ||
                            (i > 1 && (IsCol(a[i - 1]) && (IsCol(a[i + 1])))) ||
                            (i > 1 && (IsRow(a[i - 1]) && (IsRow(a[i + 1])))))
                        {
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                bool ret;
                if (r1 != "" && c1 != "" && r2 == "" && c2 == "")   //Single Cell
                {
                    var column = GetColumn(c1);
                    var row = int.Parse(r1);
                    ret = (column >= 1 && column <= ExcelPackage.MaxColumns && row >= 1 && row <= ExcelPackage.MaxRows);
                }
                else if (r1 != "" && r2 != "" && c1 != "" && c2 != "") //Range
                {
                    var iR1 = int.Parse(r1);
                    var iC1 = GetColumn(c1);
                    var iR2 = int.Parse(r2);
                    var iC2 = GetColumn(c2);

                    ret = iC1 <= iC2 && iR1 <= iR2 &&
                        iC1 >= 1 && iC2 <= ExcelPackage.MaxColumns &&
                        iR1 >= 1 && iR2 <= ExcelPackage.MaxRows;

                }
                else if (r1 == "" && r2 == "" && c1 != "" && c2 != "") //Full Column
                {
                    var iC1 = GetColumn(c1);
                    var iC2 = GetColumn(c2);
                    ret = iC1 <= iC2 &&
                        iC1 >= 1 && iC2 <= ExcelPackage.MaxColumns;
                }
                else if (r1 != "" && r2 != "" && c1 == "" && c2 == "")
                {
                    var iR1 = int.Parse(r2);
                    var iR2 = int.Parse(r2);

                    ret = int.Parse(r1) <= iR2 &&
                        iR1 >= 1 &&
                        iR2 <= ExcelPackage.MaxRows;
                }
                else
                {
                    return false;
                }
                if (ret == false) return false;
            }
            return true;
        }

        private static bool IsCol(char c)
        {
            return c >= 'A' && c <= 'Z';
        }
        private static bool IsRow(char r)
        {
            return r >= '0' && r <= '9';
        }

        /// <summary>
        /// Checks that a cell address (e.g. A5) is valid.
        /// </summary>
        /// <param name="cellAddress">The alphanumeric cell address</param>
        /// <returns>True if the cell address is valid</returns>
        public static bool IsValidCellAddress(string cellAddress)
        {
            bool result = false;
            try
            {
                int row, col;
                if (GetRowColFromAddress(cellAddress, out row, out col))
                {
                    if (row > 0 && col > 0 && row <= ExcelPackage.MaxRows && col <= ExcelPackage.MaxColumns)
                        result = true;
                    else
                        result = false;
                }
            }
            catch { }
            return result;
        }
        #endregion
        #region UpdateFormulaReferences
        /// <summary>
        /// Updates the Excel formula so that all the cellAddresses are incremented by the row and column increments
        /// if they fall after the afterRow and afterColumn.
        /// Supports inserting rows and columns into existing templates.
        /// </summary>
        /// <param name="formula">The Excel formula</param>
        /// <param name="rowIncrement">The amount to increment the cell reference by</param>
        /// <param name="colIncrement">The amount to increment the cell reference by</param>
        /// <param name="afterRow">Only change rows after this row</param>
        /// <param name="afterColumn">Only change columns after this column</param>
        /// <param name="currentSheet">The sheet that contains the formula currently being processed.</param>
        /// <param name="modifiedSheet">The sheet where cells are being inserted or deleted.</param>
        /// <param name="setFixed">Fixed address</param>
        /// <returns>The updated version of the <paramref name="formula"/>.</returns>
        internal static string UpdateFormulaReferences(string formula, int rowIncrement, int colIncrement, int afterRow, int afterColumn, string currentSheet, string modifiedSheet, bool setFixed = false)
        {
            try
            {
                var sct = new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty);
                var tokens = sct.Tokenize(formula);
                var f = "";
                foreach (var t in tokens)
                {
                    if (t.TokenTypeIsSet(TokenType.ExcelAddress))
                    {
                        var address = new ExcelAddressBase(t.Value);

                        if ((!string.IsNullOrEmpty(address._wb) || !IsReferencesModifiedWorksheet(currentSheet, modifiedSheet, address)) && !setFixed)
                        {
                            f += address.Address;
                            continue;
                        }

                        if (!string.IsNullOrEmpty(address._ws)) //The address has worksheet.
                        {
                            if(t.Value.IndexOf("'!")>=0)
                            {
                                f += $"'{address._ws}'!";
                            }
                            else
                            {
                                f += $"{address._ws}!";
                            }
                        }
                        if (!address.IsFullColumn)
                        {
                            if (rowIncrement > 0)
                            {
                                address = address.AddRow(afterRow, rowIncrement, setFixed);
                            }
                            else if (rowIncrement < 0)
                            {
                                if(address._fromRowFixed==false && (address._fromRow>=afterRow && address._toRow<afterRow-rowIncrement))
                                {
                                    address=null;
                                }
                                else
                                {
                                    address = address.DeleteRow(afterRow, -rowIncrement, setFixed);
                                }
                            }
                        }
                        if (address!=null && !address.IsFullRow)
                        {
                            if (colIncrement > 0)
                            {
                                address = address.AddColumn(afterColumn, colIncrement, setFixed);
                            }
                            else if (colIncrement < 0)
                            {
                                if (address._fromColFixed == false && (address._fromCol >= afterColumn && address._toCol < afterColumn - colIncrement))
                                {
                                    address = null;
                                }
                                else
                                {
                                    address = address.DeleteColumn(afterColumn, -colIncrement, setFixed);
                                }
                            }
                        }

                        if (address == null || !address.IsValidRowCol())
                        {
                            f += "#REF!";
                        }
                        else
                        {
                            var ix = address.Address.LastIndexOf('!');
                            if (ix > 0)
                            {
                                f += address.Address.Substring(ix + 1);
                            }
                            else
                            {
                                f += address.Address;
                            }
                        }


                    }
                    else
                    {
                        f += t.Value;
                    }
                }
                return f;
            }
            catch //Invalid formula, return formula
            {
                return formula;
            }
        }

        private static bool IsReferencesModifiedWorksheet(string currentSheet, string modifiedSheet, ExcelAddressBase a)
        {
            return (string.IsNullOrEmpty(a._ws) && currentSheet.Equals(modifiedSheet, StringComparison.CurrentCultureIgnoreCase)) ||
                                         modifiedSheet.Equals(a._ws, StringComparison.CurrentCultureIgnoreCase);
        }

        /// <summary>
        /// Updates all formulas after a worksheet has been renamed
        /// </summary>
        /// <param name="formula">The formula to be updated.</param>
        /// <param name="oldName">The old sheet name.</param>
        /// <param name="newName">The new sheet name.</param>
        /// <returns>The formula to be updated.</returns>
        internal static string UpdateSheetNameInFormula(string formula, string oldName, string newName)
        {
            if (string.IsNullOrEmpty(oldName) || string.IsNullOrEmpty(newName))
                throw new ArgumentNullException("Sheet name can't be empty");

            try
            {
                var sct = new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty);
                var retFormula = "";
                foreach (var token in sct.Tokenize(formula))
                {
                    if (token.TokenTypeIsSet(TokenType.ExcelAddress)) //Address
                    {
                        var address = new ExcelAddressBase(token.Value);
                        if (address == null || !address.IsValidRowCol())
                        {
                            retFormula += "#REF!";
                        }
                        else
                        {
                            address.ChangeWorksheet(oldName, newName);
                            retFormula += address.Address;
                        }
                    }
                    else
                    {
                        retFormula += token.Value;
                    }
                }
                return retFormula;
            }
            catch //if we have an exception, return the original formula.
            {
                return formula;
            }
        }
        #endregion
        #endregion
        #endregion
    }
}
