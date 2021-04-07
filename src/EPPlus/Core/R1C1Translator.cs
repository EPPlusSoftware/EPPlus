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
using System;
using System.Linq;

namespace OfficeOpenXml.Core
{
    /// <summary>
    /// Translate addresses between the R1C1 and A1 notation
    /// </summary>
    public static class R1C1Translator
    {
        private struct R1C1
        {
            public bool hasRow;
            public bool hasCol;
            public int Row;
            public int Col ;
            public int RowOffset;
            public int ColOffset;
        }
        /// <summary>
        /// Translate addresses in a formula from R1C1 to A1
        /// </summary>
        /// <param name="formula">The formula</param>
        /// <param name="row">The row of the cell to calculate from</param>
        /// <param name="col">The column of the cell to calculate from</param>
        /// <returns>The formula in A1 notation</returns>
        public static string FromR1C1Formula(string formula, int row, int col)
        {
            var lexer = new Lexer(SourceCodeTokenizer.R1C1, new SyntacticAnalyzer());
            var tokens = lexer.Tokenize(formula, null).ToArray();
            for(var ix = 0; ix < tokens.Length; ix++)
            {
                var token = tokens[ix];
                if (token.TokenTypeIsSet(TokenType.ExcelAddress) /*|| token.TokenTypeIsSet(TokenType.NameValue)*/ || token.TokenTypeIsSet(TokenType.ExcelAddressR1C1))
                {
                    var part = FromR1C1(token.Value, row, col);
                    tokens[ix] = tokens[ix].CloneWithNewValue(part);
                }

            }
            var ret = string.Join("", tokens.Select(x => x.TokenTypeIsSet(TokenType.StringContent) ? x.Value.Replace("\"", "\"\"") :  x.Value).ToArray());
            return ret;
        }
        /// <summary>
        /// Translate addresses in a formula from A1 to R1C1
        /// </summary>
        /// <param name="formula">The formula</param>
        /// <param name="row">The row of the cell to calculate from</param>
        /// <param name="col">The column of the cell to calculate from</param>
        /// <returns>The formula in R1C1 notation</returns>        
        public static string ToR1C1Formula(string formula, int row, int col)
        {
            var lexer = new Lexer(SourceCodeTokenizer.Default, new SyntacticAnalyzer());
            var tokens = lexer.Tokenize(formula, null).ToArray();
            for (var ix = 0; ix < tokens.Length; ix++)
            {
                var token = tokens[ix];
                if (token.TokenTypeIsSet(TokenType.ExcelAddress) || token.TokenTypeIsSet(TokenType.ExcelAddressR1C1))
                {
                    var part = ToR1C1(new ExcelAddressBase(token.Value), row, col);
                    tokens[ix] = tokens[ix].CloneWithNewValue(part);
                }

            }
            var ret = string.Join("", tokens.Select(x => x.TokenTypeIsSet(TokenType.StringContent) ? x.Value.Replace("\"", "\"\"") : x.Value).ToArray());
            return ret;
        }
        /// <summary>
        /// Translate an address from R1C1 to A1
        /// </summary>
        /// <param name="r1C1Address">The address</param>
        /// <param name="row">The row of the cell to calculate from</param>
        /// <param name="col">The column of the cell to calculate from</param>
        /// <returns>The address in A1 notation</returns>        
        public static string FromR1C1(string r1C1Address, int row, int col)
        {
            if (ExcelAddress.IsTableAddress(r1C1Address)) return r1C1Address;
            var addresses = ExcelAddressBase.SplitFullAddress(r1C1Address);
            var ret = "";
            foreach(var address in addresses)
            {
                ret += ExcelCellBase.GetFullAddress(address[0], address[1], FromR1C1SingleAddress(address[2], row, col))+",";
            }
            return ret.Length==0?"":ret.Substring(0,ret.Length-1);
        }

        private static string FromR1C1SingleAddress(string r1C1Address, int row, int col)
        {
            R1C1 firstCell = new R1C1();
            var currentCell = firstCell;
            bool isRow = false;
            bool isSecond = false;
            string num = "";
            for (int i = 0; i < r1C1Address.Length; i++)
            {
                switch (r1C1Address[i])
                {
                    case 'R':
                    case 'r':
                        currentCell.hasRow = true;
                        isRow = true;
                        break;
                    case 'C':
                    case 'c':
                        if (!string.IsNullOrEmpty(num))
                        {
                            currentCell.Row = int.Parse(num);
                            num = "";
                        }
                        currentCell.hasCol = true;
                        isRow = false;
                        break;
                    case ':':
                        if (!string.IsNullOrEmpty(num))
                        {
                            if (isRow)
                            {
                                currentCell.Row = int.Parse(num);
                            }
                            else
                            {
                                currentCell.Col = int.Parse(num);
                            }
                            num = "";
                        }
                        firstCell = currentCell;
                        currentCell = new R1C1();
                        isSecond = true;
                        isRow = false;
                        break;
                    case '[':
                        break;
                    case ']':
                        if (isRow)
                        {
                            currentCell.RowOffset = int.Parse(num);
                        }
                        else
                        {
                            currentCell.ColOffset = int.Parse(num);
                        }
                        num = "";
                        break;
                    default:
                        if ((r1C1Address[i] >= '0' && r1C1Address[i] <= '9') || r1C1Address[i] == '-' || r1C1Address[i] == '+')
                            num += r1C1Address[i];
                        else
                            return r1C1Address; //This is not a R1C1 Address. Return the address without any change.
                        break;
                }
            }
            if (!string.IsNullOrEmpty(num))
            {
                if (isRow)
                {
                    currentCell.Row = int.Parse(num);
                }
                else
                {
                    currentCell.Col = int.Parse(num);
                }
            }

            if (isSecond == false)
            {
                if (currentCell.hasRow == false || currentCell.hasCol == false)
                {
                    var cell = GetCell(currentCell, row, col);
                    return $"{cell}:{cell}";
                }
                else
                {
                    return GetCell(currentCell, row, col);
                }
            }
            else
            {
                var cell1 = GetCell(firstCell, row, col);
                var cell2 = GetCell(currentCell, row, col);
                if (cell1 == cell2)
                    return cell1;
                else
                    return $"{cell1}:{cell2}";
            }
        }

        /// <summary>
        /// Translate an address from A1 to R1C1
        /// </summary>
        /// <param name="address">The address</param>
        /// <param name="row">The row of the cell to calculate from</param>
        /// <param name="col">The column of the cell to calculate from</param>
        /// <returns>The address in R1C1 notation</returns>        
        public static string ToR1C1(ExcelAddressBase address, int row, int col)
        {
            string returnAddress;
            if(address.IsFullRow) //Full Row
            {
                if(address._fromRow==address._toRow && address._fromRowFixed == address._toRowFixed)
                {
                    returnAddress=GetCellAddress("R",address._fromRow, row, address._fromRowFixed);
                }
                else
                {
                    returnAddress = GetCellAddress("R", address._fromRow, row, address._fromRowFixed) + ":" + GetCellAddress("R", address._toRow, row, address._toRowFixed);
                }
            }
            else if(address.IsFullColumn) //Full Column
            {
                if (address._fromCol == address._toCol && address._fromColFixed == address._toColFixed)
                {
                    returnAddress = GetCellAddress("C", address._fromCol, col, address._fromColFixed);
                }
                else
                {
                    returnAddress = GetCellAddress("C", address._fromCol, col, address._fromColFixed) + ":" + GetCellAddress("C", address._toCol, col, address._toColFixed);
                }

            }
            else if(address.Table!=null)
            {
                return address.Address;
            }
            else
            {
                if (address.IsSingleCell)
                {
                    returnAddress = GetCellAddress("R", address._fromRow, row, address._fromRowFixed) + GetCellAddress("C", address._fromCol, col, address._fromColFixed);
                }
                else
                {
                    returnAddress = GetCellAddress("R", address._fromRow, row, address._fromRowFixed) + GetCellAddress("C", address._fromCol, col, address._fromColFixed) + ":" +
                           GetCellAddress("R", address._toRow, row, address._toRowFixed) + GetCellAddress("C", address._toCol, col, address._toColFixed);
                }
            }
            return ExcelAddressBase.GetFullAddress(address._wb, address._ws, returnAddress);
        }

        private static string GetCellAddress(string RC, int fromRow, int row, bool isFixed)
        {
            if (isFixed)
            {
                return $"{RC}{fromRow}";
            }
            else
            {
                if (fromRow == row)
                {
                    return RC;

                }
                else
                {
                    return $"{RC}[{fromRow - row}]";
                }
            }
        }

        private static string GetCell(R1C1 currentCell, int refRow, int refCol)
        {
            string ret="";

            if (currentCell.hasCol)
            {
                if (currentCell.Col > 0)
                {
                    ret = $"${ExcelCellBase.GetColumnLetter(currentCell.Col)}";
                }
                else
                {
                    if (refCol + currentCell.ColOffset < 1) return "#REF!";
                    ret = ExcelCellBase.GetColumnLetter(refCol + currentCell.ColOffset);
                }
            }

            if (currentCell.hasRow)
            {
                if (currentCell.Row > 0)
                {
                    ret += $"${currentCell.Row}";
                }
                else
                {
                    if (refRow + currentCell.RowOffset < 1) return "#REF!";
                    ret += (refRow + currentCell.RowOffset).ToString();
                }
            }
            return ret;
        }
    }
}
