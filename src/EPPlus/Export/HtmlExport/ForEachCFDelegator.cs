using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class ForEachCFDelegator : Dictionary<string, List<ExcelConditionalFormattingRule>>
    {
        internal ForEachCFDelegator(Dictionary<string, List<ExcelConditionalFormattingRule>> init)
            : base(init) { }

        internal void FuncOnEachElement<T>(string address, Func<ExcelConditionalFormattingRule, T> function)
        {
            if (address != null && ContainsKey(address))
            {
                foreach (var cf in this[address])
                {
                    function(cf);
                }
            }
        }
    }
}
