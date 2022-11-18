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

        internal static string ShiftAddressRowsInFormula(string worksheetName, string formula, int currentRow, int newRow)
        {
            if (string.IsNullOrEmpty(formula)) return formula;
            var tokens = SourceCodeTokenizer.Default.Tokenize(formula, worksheetName);
            if (!tokens.Any(x => x.TokenTypeIsSet(TokenType.ExcelAddress))) return formula;
            var resultTokens = new List<Token>();
            foreach (var token in tokens)
            {
                if (!token.TokenTypeIsSet(TokenType.ExcelAddress))
                {
                    resultTokens.Add(token);
                }
                else
                {
                    var addresses = new List<ExcelCellAddress>();
                    var adr = new ExcelAddressBase(token.Value);
                    // if the formula is a table formula (relative) keep it as it is
                    if (adr.Table == null)
                    {
                        var newAdr = adr.AddRow(currentRow, newRow, true);
                        var newToken = new Token(newAdr.FullAddress, TokenType.ExcelAddress);
                        resultTokens.Add(newToken);
                    }
                    else
                    {
                        resultTokens.Add(token);
                    }
                }
            }
            var result = new StringBuilder();
            foreach (var token in resultTokens)
            {
                result.Append(token.Value);
            }
            return result.ToString();
        }

        internal static string ShiftAddressColumnsInFormula(string worksheetName, string formula, int currentColumn, int newColumn)
        {
            if (string.IsNullOrEmpty(formula)) return formula;
            var tokens = SourceCodeTokenizer.Default.Tokenize(formula, worksheetName);
            if (!tokens.Any(x => x.TokenTypeIsSet(TokenType.ExcelAddress))) return formula;
            var resultTokens = new List<Token>();
            foreach (var token in tokens)
            {
                if (!token.TokenTypeIsSet(TokenType.ExcelAddress))
                {
                    resultTokens.Add(token);
                }
                else
                {
                    var addresses = new List<ExcelCellAddress>();
                    var adr = new ExcelAddressBase(token.Value);
                    // if the formula is a table formula (relative) keep it as it is
                    if (adr.Table == null)
                    {
                        var newAdr = adr.AddColumn(currentColumn, newColumn, true);
                        var newToken = new Token(newAdr.FullAddress, TokenType.ExcelAddress);
                        resultTokens.Add(newToken);
                    }
                    else
                    {
                        resultTokens.Add(token);
                    }
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
