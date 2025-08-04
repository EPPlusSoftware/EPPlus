using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Utils.Formula
{
    internal static class FormulaUtils
    {
        /// <summary>
        /// This method detects cell addresses and updates them so they have reference to the supplied worksheet. Example: SUM(A1 + B2) becomes SUM('Sheet 1'A$1$ + 'Sheet 1'!$B$2)
        /// </summary>
        /// <param name="formula">Formula to check</param>
        /// <param name="ws">Worksheet to add</param>
        /// <param name="allowRelativeAddress">If address is relative or absolute. Default is absolute.</param>
        /// <returns>A string containing the formula with worksheet reference on cell addresses.</returns>
        /// <exception cref="InvalidFormulaException">If worksheet is null and a cell address is found this method will throw this exception. Setting worksheet to null could be useful for checking validity of formulas on workbook.</exception>
        internal static string AddWorksheetReferenceToFormula(string formula, ExcelWorksheet ws, bool allowRelativeAddress = false)
        {
            bool isWsNull = ws == null ? true : false;
            var tokens = SourceCodeTokenizer.Default.Tokenize(formula);
            Dictionary<int, Token> addresses = new Dictionary<int, Token>();
            List<Token> fixedTokens = new List<Token>();
            //Collect tokens
            for (int i = 0; i < tokens.Count; i++)
            {
                if (tokens[i].TokenType==TokenType.WorksheetNameContent)
                {
                    isWsNull = false;
                }
                else if(tokens[i].TokenType == TokenType.Operator && tokens[i].Value!=":")
                {
                    isWsNull = ws == null ? true : false;
                }
                // this should be fixed more permanently. for now we just avoid crash due to the InvalidFormulaException below.
                if (tokens[i].TokenType == TokenType.ExternalReference)
                {
                    return formula;
                }
                if (tokens[i].TokenType == TokenType.CellAddress)
                {
                    if (isWsNull) throw new InvalidFormulaException("Formula with cell address must have a worksheet.");
                    if (i == 0)
                    {
                        addresses.Add(i, tokens[i]);
                    }
                    else if (tokens[i - 1].TokenType != TokenType.WorksheetName)
                    {
                        addresses.Add(i, tokens[i]);
                    }
                }
            }
            //if no cell address tokens found we can quit early and return the formula as is.
            if (addresses.Count == 0)
            {
                return formula;
            }
            //Update tokens
            for (int i = 0; i < tokens.Count; i++)
            {
                if (addresses.ContainsKey(i))
                {
                    string fullAddress = allowRelativeAddress ? ws.Cells[addresses[i].Value].FullAddress : ws.Cells[addresses[i].Value].FullAddressAbsolute;
                    var addressTokens = SourceCodeTokenizer.Default.Tokenize(fullAddress);
                    for (int j = 0; j < addressTokens.Count; j++)
                    {
                        fixedTokens.Add(addressTokens[j]);
                    }
                }
                else
                {
                    fixedTokens.Add(tokens[i]);
                }
            }
            //Create new formula string.
            return string.Concat(fixedTokens.Select(t => t.Value));
        }

    }
}
