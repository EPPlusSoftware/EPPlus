using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Export.ToCollection
{
    internal class ToCollectionRange
    {
        internal static List<string> GetRangeHeaders(ExcelRangeBase range, string[] headers, int? headerRow)
        {
            List<string> headersList;
            if (headers == null || headers.Length == 0)
            {
                headersList = new List<string>();
                if (headerRow.HasValue == false) return headersList;

                for (int c = range._fromCol; c <= range._toCol; c++)
                {
                    var h = range.Worksheet.Cells[range._fromRow + headerRow.Value, c].Text;
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
                headersList = new List<string>(headers);
            }

            return headersList;
        }
    }

}

