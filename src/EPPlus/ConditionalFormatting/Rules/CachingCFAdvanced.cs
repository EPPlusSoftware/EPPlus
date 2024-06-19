using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting.Rules
{
    internal class CachingCFAdvanced: CachingCF
    {
        protected double _highest = double.NaN;
        protected double _lowest = double.NaN;

        internal CachingCFAdvanced(CachingCF copy, ExcelWorksheet newWs) : base(copy, newWs)
        {
        }

        internal CachingCFAdvanced(eExcelConditionalFormattingRuleType cfType, ExcelAddress address, int priority, ExcelWorksheet worksheet) : base(cfType, address, priority, worksheet)
        {
        }

        internal CachingCFAdvanced(eExcelConditionalFormattingRuleType cfType, ExcelAddress address, ExcelWorksheet ws, XmlReader xr) : base(cfType, address, ws, xr)
        {
        }

        protected override void UpdateCellValueCache(bool asStrings = false, bool cacheOnlyNumeric = false)
        {
            base.UpdateCellValueCache();
            var values = cellValueCache.OrderBy(n => n);
            _highest = Convert.ToDouble(values.Last());
            _lowest = Convert.ToDouble(values.First());
        }
    }
}
