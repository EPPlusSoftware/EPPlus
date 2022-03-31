using OfficeOpenXml.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class CssRangeExporterBase : AbstractExporter
    {
        public CssRangeExporterBase(Dictionary<string, int> styleCache, List<string> dataTypes, HtmlRangeExportSettings settings, ExcelRangeBase[] ranges)
        {
            _styleCache = styleCache;
            _dataTypes = dataTypes;
            Settings = settings;
        }

        protected Dictionary<string, int> _styleCache;
        protected List<string> _dataTypes;
        protected HtmlRangeExportSettings Settings;
        protected EPPlusReadOnlyList<ExcelRangeBase> _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

        private void AddRange(ExcelRangeBase range)
        {
            if (range.IsFullColumn && range.IsFullRow)
            {
                _ranges.Add(new ExcelRangeBase(range.Worksheet, range.Worksheet.Dimension.Address));
            }
            else
            {
                _ranges.Add(range);
            }
        }
    }
}
