using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;

namespace OfficeOpenXml.Table
{
    internal class TableAdjustFormula
    {
        ExcelTable _tbl;
        public TableAdjustFormula(ExcelTable tbl)
        {
            _tbl = tbl;
        }

        internal void AdjustFormulas(string prevName, string name)
        {
            foreach (var ws in _tbl.WorkSheet.Workbook.Worksheets)
            {
                foreach (var tbl in ws.Tables)
                {
                    foreach (var c in tbl.Columns)
                    {
                        if (!string.IsNullOrEmpty(c.CalculatedColumnFormula))
                        {
                            c.CalculatedColumnFormula = ReplaceTableName(c.CalculatedColumnFormula, prevName, name);
                        }
                    }
                }

                var cse = new CellStoreEnumerator<object>(ws._formulas);
                while (cse.Next())
                {
                    if (cse.Value is string f)
                    {
                        if (f.IndexOf(prevName, StringComparison.InvariantCultureIgnoreCase) > -1)
                        {
                            ws._formulas.SetValue(cse.Row, cse.Column, ReplaceTableName(f, prevName, name));
                        }
                    }
                }

                foreach (var sf in ws._sharedFormulas.Values)
                {
                    if (sf.Formula.IndexOf(prevName, StringComparison.InvariantCultureIgnoreCase) > -1)
                    {
                        sf.Formula = ReplaceTableName(sf.Formula, prevName, name);
                    }
                }

                foreach (var n in ws.Names)
                {
                    AdjustName(n, prevName, name);
                }
            }

            foreach (var n in _tbl.WorkSheet.Workbook.Names)
            {
                AdjustName(n, prevName, name);
            }
        }

        private void AdjustName(ExcelNamedRange n, string prevName, string name)
        {
            if (!string.IsNullOrEmpty(n.Formula))
            {
                if (n.Formula.IndexOf(prevName, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    n.Formula = ReplaceTableName(n.Formula, prevName, name);
                }
            }
            else if (n.IsName == false)
            {
                if (n.Address.IndexOf(prevName, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    n.Address = ReplaceTableName(n.Address, prevName, name);
                }
            }
        }

        private string ReplaceTableName(string formula, string prevName, string name)
        {
            var tokens = _tbl.WorkSheet.Workbook.FormulaParser.Tokenizer.Tokenize(formula);
            var f = "";
            foreach (var t in tokens)
            {
                if (t.TokenTypeIsSet(TokenType.TableName) && t.Value.Equals(prevName))
                {
                    f += name;
                }
                else
                {
                    f += t.Value;
                }
            }

            return f;
        }
    }
}

