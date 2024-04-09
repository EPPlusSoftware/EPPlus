using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Export.ToCollection.Exceptions;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.ToCollection
{
    internal class ToCollectionRange
    {
        internal static List<string> GetRangeHeaders(ExcelRangeBase range, string[] headers, int? headerRow, ToCollectionRangeOptions options)
        {
            var start1 = range._fromCol;
            var end1 = range._toCol;
            var start2 = range._fromRow;
            var end2 = range._toRow;
            if (options.DataIsTransposed)
            {
                start2 = range._fromCol;
                end2 = range._toCol;
                start1 = range._fromRow;
                end1 = range._toRow;
            }

            List<string> headersList;
            if (headers == null || headers.Length == 0)
            {
                headersList = new List<string>();
                if (headerRow.HasValue == false) return headersList;

                for (int c = start1; c <= end1; c++)
                {
                    var h = options.DataIsTransposed ? range.Worksheet.Cells[c, start2 + headerRow.Value].Text : range.Worksheet.Cells[start2 + headerRow.Value, c].Text;
                    if (string.IsNullOrEmpty(h))
                    {
                        throw new InvalidOperationException("Header cells cannot be empty");
                    }
                    if (headersList.Contains(h))
                    {
                        throw new InvalidOperationException($"Header cells must be unique. Value : {h}");
                    }
                    headersList.Add(h);
                }
            }
            else
            {
                if(headers.Length > range.Columns)
                {
                    throw new InvalidOperationException("ToCollectionOptions.Headers[] contain more items than the columns in the range.");
                }
                headersList = new List<string>(headers);
            }

            return headersList;
        }
        internal static List<T> ToCollection<T>(ExcelRangeBase range, Func<ToCollectionRow, T> setRow, ToCollectionRangeOptions options)
        {
            var ret = new List<T>();
            var start1 = range._fromCol;
            var end1 = range._toCol;
            var start2 = range._fromRow;
            var end2 = range._toRow;
            if (options.DataIsTransposed)
            {
                start2 = range._fromCol;
                end2 = range._toCol;
                start1 = range._fromRow;
                end1 = range._toRow;
            }
            if (end2 < start2) return null;

            var headers = GetRangeHeaders(range, options.Headers, options.HeaderRow, options);

            var values = new List<ExcelValue>();
            var row = new ToCollectionRow(headers, range._workbook, options.ConversionFailureStrategy);
            var startRow = options.DataStartRow ?? ((options.HeaderRow ?? -1) + 1);
            for (int r = start2 + startRow; r <= end2; r++)
            {
                for (int c = start1; c <= end1; c++)
                {
                    if ((options.DataIsTransposed))
                    {
                        values.Add(range.Worksheet.GetCoreValueInner(c, r));
                    }
                    else
                    {
                        values.Add(range.Worksheet.GetCoreValueInner(r, c));
                    }
                }
                row._cellValues = values;
                var item = setRow(row);
                if (item != null)
                {
                    ret.Add(item);
                }

                values.Clear();
            }
            return ret;
        }
        
        internal static List<T> ToCollection<T>(ExcelRangeBase range, ToCollectionRangeOptions options)
        {
            var t = typeof(T);
            var h = GetRangeHeaders(range, options.Headers, options.HeaderRow, options);
            if (h.Count <= 0) throw new InvalidOperationException("No headers specified. Please set a ToCollectionOptions.HeaderRow or ToCollectionOptions.Headers[].");
            var mappings = ToCollectionAutomap.GetAutomapList<T>(h);
            var l = new List<T>();
            var values = new List<ExcelValue>();
            var startRow = options.DataStartRow ?? ((options.HeaderRow ?? -1) + 1);
            for (int r = range._fromRow + startRow; r <= range._toRow; r++)
            {
                var item = (T)Activator.CreateInstance(t);
                foreach (var m in mappings)
                {
                    var v = range.Worksheet.GetValueInner(r, m.Index + range._fromCol);
                    try
                    {
                        m.PropertyInfo.SetValue(item, v, null);
                    }
                    catch (Exception ex)
                    {
                        if (options.ConversionFailureStrategy == ToCollectionConversionFailureStrategy.Exception)
                        {
                            throw new EPPlusDataTypeConvertionException($"Failure to convert value {v} for index {m.Index}", ex);
                        }
                        else
                        {
                            m.PropertyInfo.SetValue(item, default(T), null);
                        }
                    }
                }

                l.Add(item);
            }

            return l;
        }
    }

}

