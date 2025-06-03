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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Utils
{
    /// <summary>
    /// A utility to work with Excel addresses
    /// </summary>
    public static class AddressUtility
    {
        /// <summary>
        /// Parse an entire column selection, e.g A:A
        /// </summary>
        /// <param name="address">The entire address</param>
        /// <returns></returns>
        public static string ParseEntireColumnSelections(string address)
        {
            string parsedAddress = address;
            var matches = Regex.Matches(address, "[A-Z]+:[A-Z]+");
            foreach (Match match in matches)
            {
                AddRowNumbersToEntireColumnRange(ref parsedAddress, match.Value);
            }
            return parsedAddress;
        }
        /// <summary>
        /// Add row number to entire column range
        /// </summary>
        /// <param name="address">The address</param>
        /// <param name="range">The full column range</param>
        private static void AddRowNumbersToEntireColumnRange(ref string address, string range)
        {
            var parsedRange = string.Format("{0}{1}", range, ExcelPackage.MaxRows);
            var splitArr = parsedRange.Split(new char[] { ':' });
            address = address.Replace(range, string.Format("{0}1:{1}", splitArr[0], splitArr[1]));
        }

        internal static string ShiftAddressRowsInFormula(ExcelRangeBase range, string formula, int currentRow, int rows)
        {
            if (string.IsNullOrEmpty(formula)) return formula;
            var affectedRange = new ExcelAddressBase(range.ExternalReferenceIndex, range.WorkSheetName, range._fromRow, range._fromCol, range._toRow, range._toCol);
            var tokens = SourceCodeTokenizer.Default.Tokenize(formula, affectedRange.WorkSheetName);
            if (!tokens.Any(x => x.TokenTypeIsAddress)) return formula;
            var resultTokens = new List<Token>();
            string extRef = string.Empty, ws = string.Empty;
            for(int i=0;i<tokens.Count;i++)
            {
                var token = tokens[i];
                if(token.TokenTypeIsAddressToken)
                {
                    var adr = GetAddressFromToken(tokens, ref i);
                    if (adr.IsFullColumn==false)
                    {
                        if (adr.Table == null && (adr.Collide(affectedRange) != ExcelAddressBase.eAddressCollition.No) &&
                                ((extRef == string.Empty && ws == string.Empty) ||
                                (ws.Equals(range.WorkSheetName, StringComparison.InvariantCultureIgnoreCase) && extRef == string.Empty) ||
                                 int.TryParse(extRef, out int ier) && ier == range.ExternalReferenceIndex))
                        {
                            ExcelAddressBase newAdr;
                            if(rows<0)
                            {
                                newAdr = adr.DeleteRowKeepFixed(1, Math.Abs(rows));
                            }
                            else
                            {
                                newAdr = adr.AddRow(1, rows, true);
                            }
                            if (newAdr == null)
                            {
                                resultTokens.Add(new Token("#REF!", TokenType.InvalidReference));
                            }
                            else
                            {
                                resultTokens.Add(new Token(newAdr.FullAddress, TokenType.ExcelAddress));
                            }
                        }
                        else
                        {
                            resultTokens.Add(new Token(adr.FullAddress, TokenType.ExcelAddress));
                        }
                    }
                    else
                    {
                        resultTokens.Add(new Token(adr.FullAddress, TokenType.ExcelAddress));
                    }
                }
                else
                {
                    resultTokens.Add(token);
                }
            }
            var result = new StringBuilder();
            foreach (var token in resultTokens)
            {
                result.Append(token.Value);
            }
            return result.ToString();
        }

        private static ExcelAddressBase GetAddressFromToken(IList<Token> tokens, ref int i)
        {
            var sb = new StringBuilder(tokens[i].Value);
            while (tokens.Count>i+1 && tokens[i+1].TokenTypeIsAddressToken)
            {
                sb.Append(tokens[++i].Value);
            }
            return new ExcelAddress(sb.ToString());
        }

        internal static string ShiftAddressColumnsInFormula(ExcelRangeBase range, string formula, int currentColumn, int columns)
        {
            if (string.IsNullOrEmpty(formula)) return formula;
            var affectedRange = new ExcelAddressBase(range.ExternalReferenceIndex <= 0 ? -1 : range.ExternalReferenceIndex, range.WorkSheetName, range._fromRow, range._fromCol, range._toRow, range._toCol);
            var tokens = SourceCodeTokenizer.Default.Tokenize(formula, affectedRange.WorkSheetName);
            if (tokens.Any(x => x.TokenTypeIsAddress)==false) return formula;
            var resultTokens = new List<Token>();
            string extRef = string.Empty, ws = string.Empty;
            for(var i=0;i < tokens.Count;i++)
            {
                var token = tokens[i];
                if (token.TokenTypeIsAddressToken)
                {
                    var adr = GetAddressFromToken(tokens, ref i);
                    if (adr.IsFullRow == false)
                    {
                        if (adr.Table == null && (adr.Collide(affectedRange) != ExcelAddressBase.eAddressCollition.No) &&
                                ((extRef == string.Empty && ws == string.Empty) ||
                                (ws.Equals(range.WorkSheetName, StringComparison.InvariantCultureIgnoreCase) && extRef == string.Empty) ||
                                 int.TryParse(extRef, out int ier) && ier == range.ExternalReferenceIndex))
                        {
                            ExcelAddressBase newAdr;
                            if (columns < 0)
                            {
                                newAdr = adr.DeleteColumnKeepFixed(1, Math.Abs(columns));
                            }
                            else
                            {
                                newAdr = adr.AddRow(1, columns, true);
                            }
                            if (newAdr == null)
                            {
                                resultTokens.Add(new Token("#REF!", TokenType.InvalidReference));
                            }
                            else
                            {
                                resultTokens.Add(new Token(newAdr.FullAddress, TokenType.ExcelAddress));
                            }
                        }
                        else
                        {
                            resultTokens.Add(new Token(adr.FullAddress, TokenType.ExcelAddress));
                        }
                    }
                    else
                    {
                        resultTokens.Add(new Token(adr.FullAddress, TokenType.ExcelAddress));
                    }
                }
                else
                {
                    resultTokens.Add(token);
                }
            }
            var result = new StringBuilder();
            foreach (var token in resultTokens)
            {
                result.Append(token.Value);
            }
            return result.ToString();
        }
    }
}
