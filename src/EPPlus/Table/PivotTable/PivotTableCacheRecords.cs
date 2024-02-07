/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/27/2023         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace EPPlusTest.Table.PivotTable
{
    internal class PivotTableCacheRecords 
    {
        internal PivotTableCacheRecords(PivotTableCacheInternal cache)
        {
            Cache = cache;
        }

        internal void CreateRecords()
        {
            var sr = Cache.SourceRange;
            var ws = sr.Worksheet;
            for (int i = 0; i < Cache.Fields.Count; i++)
            {
                var f = Cache.Fields[i];
                var l = new List<object>();
                var c = sr._fromCol + f.Index;
                if (f.IsRowColumnOrPage)
                {
                    f._fieldRecordIndex = new Dictionary<int, List<int>>();
					if (f.Grouping == null)
                    {
						for (int r = sr._fromRow + 1; r <= sr._toRow; r++) //Skip first row as it contains the headers.
						{
							var ix = f._cacheLookup[ws.GetValue(r, c)];
							l.Add(ix);
							var ciIx = r - (sr._fromRow + 1);
							if (f._fieldRecordIndex.ContainsKey(ix))
							{
								f._fieldRecordIndex[ix].Add(ciIx);
							}
							else
							{
								f._fieldRecordIndex.Add(ix, new List<int> { ciIx });
							}
						}
					}
                    else
                    {
						for (int r = sr._fromRow + 1; r <= sr._toRow; r++) //Skip first row as it contains the headers.
						{
							var ix = GetGroupIndex(f, ws.GetValue(r, c));

							l.Add(ix);
							var ciIx = r - (sr._fromRow + 1);
							if (f._fieldRecordIndex.ContainsKey(ix))
							{
								f._fieldRecordIndex[ix].Add(ciIx);
							}
							else
							{
								f._fieldRecordIndex.Add(ix, new List<int> { ciIx });
							}
						}
					}
				}

                else
                {
                    for (int r = sr._fromRow + 1; r <= sr._toRow; r++)
                    {
                        l.Add(ws.GetValue(r, c));
                    }
                }
                CacheItems.Add(l);
            }
        }

		private int GetGroupIndex(ExcelPivotTableCacheField f, object value)
		{
			if(f.Grouping is ExcelPivotTableFieldDateGroup dg)
			{
				return GetDateGroupIndex(dg, value);
			}
			else if(f.Grouping is ExcelPivotTableFieldNumericGroup ng)
			{
				return GetNumericGroupIndex(ng, value);
			}
			return 0;
		}

		private int GetNumericGroupIndex(ExcelPivotTableFieldNumericGroup ng, object value)
		{
			if(ConvertUtil.IsNumeric(value))
			{
				var d = ConvertUtil.GetValueDouble(value);
				return (int)((d - ng.Start) / ng.Interval);
			}
			return 0;
		}

		private static int GetDateGroupIndex(ExcelPivotTableFieldDateGroup dg, object value)
		{
			var startDate = dg.StartDate ?? DateTime.MinValue;
			var dtNull = ConvertUtil.GetValueDate(value);
			if (dtNull == null) return -1;
			var dt=dtNull.Value;
			switch (dg.GroupBy)
			{
				case eDateGroupBy.Years:
					return dt.Year - startDate.Year;
				case eDateGroupBy.Quarters:
					return (((dt.Month - (dt.Month - 1) % 3) + 1) / 3)-1;
				case eDateGroupBy.Months:
					return dt.Month-1;
				case eDateGroupBy.Days:
					return GetDayGroupIndex(dg, startDate, dt);
				case eDateGroupBy.Hours:
					return dt.Hour - 1;
				case eDateGroupBy.Minutes:
					return dt.Minute - 1;
				case eDateGroupBy.Seconds:
					return dt.Second - 1;
			}
			return -1;
		}

		private static int GetDayGroupIndex(ExcelPivotTableFieldDateGroup dg, DateTime startDate, DateTime dt)
		{
			if (dt < startDate)
			{
				return 0;
			}
			else
			{
				if ((dg.GroupInterval ?? 1) == 1)
				{
					var startOfYear = new DateTime(dt.Year, 1, 1);
					if (DateTime.IsLeapYear(dt.Year))
					{
						return (dt - startOfYear).Days + 1;
					}
					else
					{
						if (dt.Month < 3)
						{
							return (dt - startOfYear).Days + 1;
						}
						else
						{
							return (dt - startOfYear).Days + 2; //Series is leap year, so add one extra if after last of feb.
						}
					}
				}
				else
				{
					return (int)((dt - dg.StartDate.Value).Days / dg.GroupInterval.Value);
				}
			}
		}

		internal PivotTableCacheInternal Cache { get; }
        public List<List<object>> CacheItems 
        { 
            get; 
        }= new List<List<object>>();
        public int RecordCount
        {
            get
            {
                if(CacheItems==null || CacheItems.Count==0 || CacheItems[0].Count==0)
                {
                    return 0;
                }
                return CacheItems[0].Count;
            }
        }
    }
}
