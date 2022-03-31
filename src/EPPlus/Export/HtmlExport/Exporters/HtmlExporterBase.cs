using OfficeOpenXml.Core;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class HtmlExporterBase : AbstractExporter
    {
        public HtmlExporterBase(ExcelRangeBase range)
        {
            Require.Argument(range).IsNotNull("range");
        _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

            if (range.Addresses == null)
            {
                AddRange(range);
            }
            else
            {
                foreach (var address in range.Addresses)
                {
                    AddRange(range.Worksheet.Cells[address.Address]);
                }
            }

            LoadRangeImages(_ranges._list);
        }

        protected List<int> _columns = new List<int>();

        private void LoadVisibleColumns(ExcelRangeBase range)
        {
            var ws = range.Worksheet;
            _columns = new List<int>();
            for (int col = range._fromCol; col <= range._toCol; col++)
            {
                var c = ws.GetColumn(col);
                if (c == null || (c.Hidden == false && c.Width > 0))
                {
                    _columns.Add(col);
                }
            }
        }

        public HtmlExporterBase(ExcelRangeBase[] ranges)
        {
            Require.Argument(ranges).IsNotNull("ranges");
            _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

            foreach (var range in ranges)
            {
                AddRange(range);
            }

            LoadRangeImages(_ranges._list);
        }

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
