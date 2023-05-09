using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeOpenXml.ExternalReferences
{
    internal static class ExternalLinksHandler
    {
        /// <summary>
        /// Clears all formulas leaving the value only for formulas containing external links
        /// </summary>
        /// <param name="wb"></param>
        internal static void BreakAllFormulaLinks(ExcelWorkbook wb)
        {
            foreach (var ws in wb.Worksheets)
            {
                var _deletedFormulas = new List<int>();
                foreach (var sh in ws._sharedFormulas.Values)
                {
                    sh.SetTokens(ws.Name);
                    if (HasFormulaExternalReference(sh.Tokens))
                    {
                        ExcelCellBase.GetRowColFromAddress(sh.Address, out int fromRow, out int fromCol, out int toRow, out int toCol);
                        ws._formulas.Clear(fromRow, fromCol, toRow - fromRow + 1, toCol - fromCol + 1);
                        ws._formulaTokens?.Clear(fromRow, fromCol, toRow - fromRow + 1, toCol - fromCol + 1);
                        _deletedFormulas.Add(sh.Index);
                    }
                }
                _deletedFormulas.ForEach(x => ws._sharedFormulas.Remove(x));

                var enumerator = new CellStoreEnumerator<object>(ws._formulas);
                foreach (var f in enumerator)
                {
                    if (f is string formula)
                    {
                        IEnumerable<Token> t = ws._formulaTokens?.GetValue(enumerator.Row, enumerator.Column);
                        if (t == null)
                        {
                            t = SourceCodeTokenizer.Default.Tokenize(formula, ws.Name);
                        }
                        if (HasFormulaExternalReference(t))
                        {
                            ws._formulas.Clear(enumerator.Row, enumerator.Column, 1, 1);
                            ws._formulaTokens?.Clear(enumerator.Row, enumerator.Column, 1, 1);
                        }
                    }
                }
                HandleNames(wb, ws.Name, ws.Names, -1);
            }
            HandleNames(wb, "", wb.Names, -1);
        }
        internal static void BreakFormulaLinks(ExcelWorkbook wb, int ix, bool delete)
        {
            foreach (var ws in wb.Worksheets)
            {
                var _deletedFormulas = new List<int>();
                foreach (var sh in ws._sharedFormulas.Values)
                {
                    sh.SetTokens(ws.Name);
                    if (HasFormulaExternalReference(wb, ix, sh.Tokens, out string newFormula, false))
                    {
                        ExcelCellBase.GetRowColFromAddress(sh.Address, out int fromRow, out int fromCol, out int toRow, out int toCol);
                        ws._formulas.Clear(fromRow, fromCol, toRow - fromRow + 1, toCol - fromCol + 1);
                        ws._formulaTokens?.Clear(fromRow, fromCol, toRow - fromRow + 1, toCol - fromCol + 1);
                        _deletedFormulas.Add(sh.Index);
                    }
                    else if (newFormula != sh.Formula)
                    {
                        sh.Tokens = null;
                        ExcelCellBase.GetRowColFromAddress(sh.Address, out int fromRow, out int fromCol, out int toRow, out int toCol);
                        ws._formulaTokens?.Clear(fromRow, fromCol, toRow - fromRow + 1, toCol - fromCol + 1);
                    }
                }

                _deletedFormulas.ForEach(x => ws._sharedFormulas.Remove(x));

                var enumerator = new CellStoreEnumerator<object>(ws._formulas);
                foreach (var f in enumerator)
                {
                    if (f is string formula)
                    {
                        IEnumerable<Token> t = ws._formulaTokens?.GetValue(enumerator.Row, enumerator.Column);
                        if (t == null)
                        {
                            t = SourceCodeTokenizer.Default.Tokenize(formula, ws.Name);
                        }
                        if (HasFormulaExternalReference(wb, ix, t, out string newFormula, false))
                        {
                            ws._formulas.Clear(enumerator.Row, enumerator.Column, 1, 1);
                            ws._formulaTokens?.Clear(enumerator.Row, enumerator.Column, 1, 1);
                        }
                        else if (newFormula != formula)
                        {
                            enumerator.Value = newFormula;
                        }
                    }
                }
                HandleNames(wb, ws.Name, ws.Names, ix);
            }

            HandleNames(wb, "", wb.Names, ix);
        }

        private static void HandleNames(ExcelWorkbook wb, string wsName, ExcelNamedRangeCollection names, int ix)
        {
            var deletedNames = new List<ExcelNamedRange>();
            foreach (var n in names)
            {
                if (string.IsNullOrEmpty(n.Formula))
                {
                    if (n.Addresses != null)
                    {
                        foreach (var a in n.Addresses)
                        {
                            if (ExcelCellBase.IsExternalAddress(a.Address))
                            {
                                var startIx = a.Address.IndexOf('[');
                                var endIx = a.Address.IndexOf(']');
                                var extRef = a.Address.Substring(startIx + 1, endIx - startIx - 1);
                                var extRefIx = wb.ExternalLinks.GetExternalLink(extRef);
                                if ((extRefIx == ix || ix==-1) && extRef!="0") //-1 means delete all external references. extRef=="0" is the current workbook
                                {
                                    //deletedNames.Add(n);
                                    n.Address = "#REF!";
                                }
                                else if (extRefIx > ix)
                                {
                                    a._address = a.Address.Substring(0, startIx+1) + (extRefIx.ToString(CultureInfo.InvariantCulture)) + a.Address.Substring(endIx);
                                }
                            }
                        }
                    }
                }
                else
                {
                    var t = SourceCodeTokenizer.Default.Tokenize(n.Formula, wsName);
                    //if (ix == -1 && HasFormulaExternalReference(t))
                    //{
                    //    //deletedNames.Add(n);
                    //}
                    //else
                    //{
                        if (HasFormulaExternalReference(wb, ix, t, out string newFormula, true))
                        {
                            //deletedNames.Add(n);
                            if(newFormula!="")
                            {
                                n.Formula = newFormula;
                            }                            
                        }
                        else if (newFormula != n.Formula)
                        {
                            n.Formula = newFormula;
                        }
                    //}
                }
            }
            //deletedNames.ForEach(x => names.Remove(x.Name));
        }
        private static bool HasFormulaExternalReference(IEnumerable<Token> tokens)
        {
            foreach (var t in tokens)
            {
                if (t.TokenTypeIsSet(FormulaParsing.LexicalAnalysis.TokenType.ExcelAddress) ||
                    t.TokenTypeIsSet(FormulaParsing.LexicalAnalysis.TokenType.NameValue) ||
                    t.TokenTypeIsSet(FormulaParsing.LexicalAnalysis.TokenType.InvalidReference))
                {
                    var address = t.Value;
                    if (address.StartsWith("[") || address.StartsWith("'["))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        private static bool HasFormulaExternalReference(ExcelWorkbook wb, int ix, IEnumerable<Token> tokens, out string newFormula, bool setRefError)
        {
            newFormula = "";
            foreach (var t in tokens)
            {
                if (t.TokenTypeIsSet(FormulaParsing.LexicalAnalysis.TokenType.ExcelAddress) ||
                    t.TokenTypeIsSet(FormulaParsing.LexicalAnalysis.TokenType.NameValue) ||
                    t.TokenTypeIsSet(FormulaParsing.LexicalAnalysis.TokenType.InvalidReference))
                {
                    var address = t.Value;
                    if (address.StartsWith("[") || address.StartsWith("'["))
                    {
                        var startIx = address.IndexOf('[');
                        var endIx = address.IndexOf(']');
                        var extRef = address.Substring(startIx+1, endIx - startIx - 1);
                        if (extRef == "0") //Current workbook
                        {
                            newFormula += address;
                        }
                        else
                        {
                            var extRefIx = wb.ExternalLinks.GetExternalLink(extRef);
                            if (extRefIx == ix || ix==-1)
                            {
                                if (setRefError)
                                {
                                    newFormula += "#REF!";
                                }
                                else
                                {
                                    return true;
                                }
                            }
                            else if (extRefIx > ix)
                            {
                                newFormula += address.Substring(0, startIx + 1) + (extRefIx.ToString(CultureInfo.InvariantCulture)) + address.Substring(endIx);
                            }
                            else
                            {
                                newFormula += address;
                            }
                        }
                    }
                    else
                    {
                        newFormula += address;
                    }
                }
                else
                {
                    newFormula += t.Value;
                }
            }
            return false;
        }
    }
}
