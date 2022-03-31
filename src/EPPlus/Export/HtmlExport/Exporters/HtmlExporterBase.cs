using OfficeOpenXml.Core;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class HtmlExporterBase : AbstractExporter
    {
        public HtmlExporterBase(Dictionary<string, int> styleCache, List<string> dataTypes, HtmlRangeExportSettings settings, ExcelRangeBase range)
        {
            _styleCache = styleCache;
            _dataTypes = dataTypes;
            Settings = settings;
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

        public HtmlExporterBase(Dictionary<string, int> styleCache, List<string> dataTypes, HtmlRangeExportSettings settings, ExcelRangeBase[] ranges)
        {
            _styleCache = styleCache;
            _dataTypes = dataTypes;
            Settings = settings;
            Require.Argument(ranges).IsNotNull("ranges");
            _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

            foreach (var range in ranges)
            {
                AddRange(range);
            }

            LoadRangeImages(_ranges._list);
        }

        protected List<int> _columns = new List<int>();
        protected Dictionary<string, int> _styleCache;
        protected List<string> _dataTypes;
        protected HtmlRangeExportSettings Settings;

        protected void LoadVisibleColumns(ExcelRangeBase range)
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

        protected string GetWorksheetClassName(string styleClassPrefix, string name, ExcelWorksheet ws, bool addWorksheetName)
        {
            if (addWorksheetName)
            {
                return styleClassPrefix + name + "-" + GetClassName(ws.Name, $"Sheet{ws.PositionId}");
            }
            else
            {
                return styleClassPrefix + name;
            }
        }

        protected string GetClassName(string className, string optionalName)
        {
            if (string.IsNullOrEmpty(optionalName)) return optionalName;

            className = className.Trim().Replace(" ", "-");
            var newClassName = "";
            for (int i = 0; i < className.Length; i++)
            {
                var c = className[i];
                if (i == 0)
                {
                    if (c == '-' || (c >= '0' && c <= '9'))
                    {
                        newClassName = "_";
                        continue;
                    }
                }

                if ((c >= '0' && c <= '9') ||
                   (c >= 'a' && c <= 'z') ||
                   (c >= 'A' && c <= 'Z') ||
                    c >= 0x00A0)
                {
                    newClassName += c;
                }
            }
            return string.IsNullOrEmpty(newClassName) ? optionalName : newClassName;
        }

        internal bool HandleHiddenRow(EpplusHtmlWriter writer, ExcelWorksheet ws, HtmlExportSettings Settings, ref int row)
        {
            if (Settings.HiddenRows != eHiddenState.Include)
            {
                var r = ws.Row(row);
                if (r.Hidden || r.Height == 0)
                {
                    if (Settings.HiddenRows == eHiddenState.IncludeButHide)
                    {
                        writer.AddAttribute("class", $"{Settings.StyleClassPrefix}hidden");
                    }
                    else
                    {
                        row++;
                        return true;
                    }
                }
            }

            return false;
        }

        internal void AddRowHeightStyle(EpplusHtmlWriter writer, ExcelRangeBase range, int row, string styleClassPrefix, bool isMultiSheet)
        {
            var r = range.Worksheet._values.GetValue(row, 0);
            if (r._value is RowInternal rowInternal)
            {
                if (rowInternal.Height != -1 && rowInternal.Height != range.Worksheet.DefaultRowHeight)
                {
                    writer.AddAttribute("style", $"height:{rowInternal.Height}pt");
                    return;
                }
            }

            var clsName = GetWorksheetClassName(styleClassPrefix, "drh", range.Worksheet, isMultiSheet);
            writer.AddAttribute("class", clsName); //Default row height
        }

        protected string GetPictureName(HtmlImage p)
        {
            var hash = ((IPictureContainer)p.Picture).ImageHash;
            var fi = new FileInfo(p.Picture.Part.Uri.OriginalString);
            var name = fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length);

            return GetClassName(name, hash);
        }
    }
}
