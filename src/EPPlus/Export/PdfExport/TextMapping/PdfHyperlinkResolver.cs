/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using System;

namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal static class PdfHyperlinkResolver
    {
        private static bool IsWebUri(string s) => s.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || s.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                                                  s.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase) || s.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase);

        internal static bool TryGetFormulaHyperlink(ExcelRangeBase cell, out string uri, out string internalRef)
        {
            uri = null;
            internalRef = null;

            var f = cell.Formula;
            if (string.IsNullOrEmpty(f)) return false;
            f = f.TrimStart();
            const string fn = "HYPERLINK(";
            if (!f.StartsWith(fn, StringComparison.OrdinalIgnoreCase)) return false;
            int i = fn.Length, start = i, depth = 0;
            bool inString = false;
            for (; i < f.Length; i++)
            {
                char c = f[i];
                if (inString)
                {
                    if (c == '"')
                    {
                        if (i + 1 < f.Length && f[i + 1] == '"') i++;
                        else inString = false;
                    }
                }
                else if (c == '"') inString = true;
                else if (c == '(') depth++;
                else if (c == ')') { if (depth == 0) break; depth--; }
                else if (c == ',' && depth == 0) break;
            }
            var arg0 = f.Substring(start, i - start).Trim();
            if (arg0.Length < 2 || arg0[0] != '"' || arg0[arg0.Length - 1] != '"')
                return false;
            var target = arg0.Substring(1, arg0.Length - 2).Replace("\"\"", "\"").Trim();
            if (target.Length == 0)
                return false;
            if (target[0] == '#')
                internalRef = target.Substring(1);
            else
                uri = NormalizeExternalTarget(target);
            return true;
        }

        internal static string NormalizeExternalTarget(string target)
        {
            if (string.IsNullOrEmpty(target)) return null;
            if (IsWebUri(target)) return target;
            if (Uri.TryCreate(target, UriKind.Absolute, out var u))
                return u.IsFile ? u.AbsoluteUri : target;
            return null;
        }

        internal static bool TryResolveInternalLink(ExcelWorksheet linkSheet, string target, out string targetSheet, out int row, out int col)
        {
            targetSheet = null; row = 0; col = 0;
            if (string.IsNullOrEmpty(target)) return false;
            var wb = linkSheet.Workbook;
            if (target.IndexOf('!') < 0)
            {
                var name = wb.Names.ContainsKey(target) ? wb.Names[target]
                         : linkSheet.Names.ContainsKey(target) ? linkSheet.Names[target] : null;
                if (name != null)
                {
                    var ws = name.Worksheet ?? linkSheet;
                    targetSheet = ws.Name; row = name.Start.Row; col = name.Start.Column;
                    return row > 0 && col > 0;
                }
            }
            var addr = new ExcelAddress(target);
            var sheet = string.IsNullOrEmpty(addr.WorkSheetName) ? linkSheet : wb.Worksheets[addr.WorkSheetName];
            if (sheet == null) return false;
            targetSheet = sheet.Name; row = addr.Start.Row; col = addr.Start.Column;
            return row > 0 && col > 0;
        }
    }
}
