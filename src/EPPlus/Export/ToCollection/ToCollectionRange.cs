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
            var fromCol = range._fromCol;
            var toCol = range._toCol;
            var fromRow = range._fromRow;
            var toRow = range._toRow;
            if (options.DataIsTransposed)
            {
                fromRow = range._fromCol;
                toRow = range._toCol;
                fromCol = range._fromRow;
                toCol = range._toRow;
            }

            List<string> headersList;
            if (headers == null || headers.Length == 0)
            {
                headersList = new List<string>();
                if (headerRow.HasValue == false) return headersList;

                for (int c = fromCol; c <= toCol; c++)
                {
                    var h = options.DataIsTransposed ? range.Worksheet.Cells[c, fromRow + headerRow.Value].Text : range.Worksheet.Cells[fromRow + headerRow.Value, c].Text;
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
            var fromCol = range._fromCol;
            var toCol = range._toCol;
            var fromRow = range._fromRow;
            var toRow = range._toRow;
            if (options.DataIsTransposed)
            {
                fromRow = range._fromCol;
                toRow = range._toCol;
                fromCol = range._fromRow;
                toCol = range._toRow;
            }
            if (toRow < fromRow) return null;

            var headers = GetRangeHeaders(range, options.Headers, options.HeaderRow, options);

            var values = new List<ExcelValue>();
            var row = new ToCollectionRow(headers, range._workbook, options.ConversionFailureStrategy);
            var startRow = options.DataStartRow ?? ((options.HeaderRow ?? -1) + 1);
            for (int r = fromRow + startRow; r <= toRow; r++)
            {
                for (int c = fromCol; c <= toCol; c++)
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

