/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/02/2020         EPPlus Software AB       EPPlus 5.4
 *************************************************************************************************/
using EPPlusTest.Table.PivotTable;
using OfficeOpenXml;
using OfficeOpenXml.Filter;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable.Filter
{
    /// <summary>
    /// Defines a pivot table filter
    /// </summary>
    public class ExcelPivotTableFilter : XmlHelper
    {
        XmlNode _filterColumnNode;
        bool _date1904;
        internal ExcelPivotTableFilter(XmlNamespaceManager nsm, XmlNode topNode, bool date1904) : base(nsm, topNode)
        {
            if (topNode.InnerXml == "")
            {
                topNode.InnerXml = "<autoFilter ref=\"A1\"><filterColumn colId=\"0\"></filterColumn></autoFilter>";
            }
            else
            {
                LoadValues();
            }

            _filterColumnNode = GetNode("d:autoFilter/d:filterColumn");
            _date1904 = date1904;

        }

        private void LoadValues()
        {
            if ((int)Type < 100) // Caption
            {
                Value1 = StringValue1;
                Value2 = StringValue2;
            }
            else
            {
                switch (Type)
                {
                    case ePivotTableFilterType.ValueEqual:
                    case ePivotTableFilterType.ValueNotEqual:
                    case ePivotTableFilterType.ValueGreaterThan:
                    case ePivotTableFilterType.ValueGreaterThanOrEqual:
                    case ePivotTableFilterType.ValueLessThan:
                    case ePivotTableFilterType.ValueLessThanOrEqual:
                        Value1 = GetXmlNodeDoubleNull("d:autoFilter/d:filterColumn/d:customFilters/d:customFilter[1]/@val");
                        break;
                    case ePivotTableFilterType.ValueBetween:
                    case ePivotTableFilterType.ValueNotBetween:
                        Value1 = GetXmlNodeDoubleNull("d:autoFilter/d:filterColumn/d:customFilters/d:customFilter[1]/@val");
                        Value2 = GetXmlNodeDoubleNull("d:autoFilter/d:filterColumn/d:customFilters/d:customFilter[2]/@val");
                        break;
                    case ePivotTableFilterType.DateEqual:
                    case ePivotTableFilterType.DateNotEqual:
                    case ePivotTableFilterType.DateNewerThan:
                    case ePivotTableFilterType.DateNewerThanOrEqual:
                    case ePivotTableFilterType.DateOlderThan:
                    case ePivotTableFilterType.DateOlderThanOrEqual:
                        Value1 = GetValueDate("d:autoFilter/d:filterColumn/d:customFilters/d:customFilter[1]/@val");
                        break;
                    case ePivotTableFilterType.DateBetween:
                    case ePivotTableFilterType.DateNotBetween:
                        Value1 = GetValueDate("d:autoFilter/d:filterColumn/d:customFilters/d:customFilter[1]/@val");
                        Value2 = GetValueDate("d:autoFilter/d:filterColumn/d:customFilters/d:customFilter[2]/@val");
                        break;
                    case ePivotTableFilterType.Count:
                    case ePivotTableFilterType.Sum:
                    case ePivotTableFilterType.Percent:
                        //<top10 val="2" top="1" percent="0" filterVal="2"/>
                        var f = new ExcelTop10FilterColumn(NameSpaceManager, GetNode("d:autoFilter/d:filterColumn"));
						Filter = f;
                        Value1 = f.Value;
						break;
				}
			}
        }

        private DateTime? GetValueDate(string xPath)
        {
            var v = GetXmlNodeDoubleNull(xPath);
            if (v.HasValue)
            {
                return DateTime.FromOADate(v.Value);
            }
            return null;
        }

        /// <summary>
        /// The id 
        /// </summary>
        public int Id
        {
            get
            {
                return GetXmlNodeInt("@id");
            }
            internal set
            {
                SetXmlNodeInt("@id", value);
            }
        }
        /// <summary>
        /// The name of the pivot filter
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
            set
            {
                SetXmlNodeString("@name", value, true);
            }
        }
        /// <summary>
        /// The description of the pivot filter
        /// </summary>
        public string Description
        {
            get
            {
                return GetXmlNodeString("@description");
            }
            set
            {
                SetXmlNodeString("@description", value, true);
            }
        }
        internal void CreateDateCustomFilter(ePivotTableDateValueFilterType type)
        {
            _filterColumnNode.InnerXml = "<customFilters/>";
            var cf = new ExcelCustomFilterColumn(NameSpaceManager, _filterColumnNode);

            eFilterOperator t;
            var v = ConvertUtil.GetValueForXml(Value1, _date1904);
            switch (type)
            {
                case ePivotTableDateValueFilterType.DateNotEqual:
                    t = eFilterOperator.NotEqual;
                    break;
                case ePivotTableDateValueFilterType.DateNewerThan:
                case ePivotTableDateValueFilterType.DateBetween:
                    t = eFilterOperator.GreaterThan;
                    break;
                case ePivotTableDateValueFilterType.DateNewerThanOrEqual:
                    t = eFilterOperator.GreaterThanOrEqual;
                    break;
                case ePivotTableDateValueFilterType.DateOlderThan:
                case ePivotTableDateValueFilterType.DateNotBetween:
                    t = eFilterOperator.LessThan;
                    break;
                case ePivotTableDateValueFilterType.DateOlderThanOrEqual:
                    t = eFilterOperator.LessThanOrEqual;
                    break;
                default:
                    t = eFilterOperator.Equal;
                    break;
            }

            var item1 = new ExcelFilterCustomItem(v, t);
            cf.Filters.Add(item1);

            if (type == ePivotTableDateValueFilterType.DateBetween)
            {
                cf.And = true;
                cf.Filters.Add(new ExcelFilterCustomItem(ConvertUtil.GetValueForXml(Value2, _date1904), eFilterOperator.LessThanOrEqual));
            }
            else if (type == ePivotTableDateValueFilterType.DateNotBetween)
            {
                cf.And = false;
                cf.Filters.Add(new ExcelFilterCustomItem(ConvertUtil.GetValueForXml(Value2, _date1904), eFilterOperator.GreaterThan));
            }
            _filter = cf;
        }

        internal void CreateDateDynamicFilter(ePivotTableDatePeriodFilterType type)
        {
            _filterColumnNode.InnerXml = "<dynamicFilter />";
            var df = new ExcelDynamicFilterColumn(NameSpaceManager, _filterColumnNode);
            switch (type)
            {
                case ePivotTableDatePeriodFilterType.LastMonth:
                    df.Type = eDynamicFilterType.LastMonth;
                    break;
                case ePivotTableDatePeriodFilterType.LastQuarter:
                    df.Type = eDynamicFilterType.LastQuarter;
                    break;
                case ePivotTableDatePeriodFilterType.LastWeek:
                    df.Type = eDynamicFilterType.LastWeek;
                    break;
                case ePivotTableDatePeriodFilterType.LastYear:
                    df.Type = eDynamicFilterType.LastYear;
                    break;
                case ePivotTableDatePeriodFilterType.M1:
                    df.Type = eDynamicFilterType.M1;
                    break;
                case ePivotTableDatePeriodFilterType.M2:
                    df.Type = eDynamicFilterType.M2;
                    break;
                case ePivotTableDatePeriodFilterType.M3:
                    df.Type = eDynamicFilterType.M3;
                    break;
                case ePivotTableDatePeriodFilterType.M4:
                    df.Type = eDynamicFilterType.M4;
                    break;
                case ePivotTableDatePeriodFilterType.M5:
                    df.Type = eDynamicFilterType.M5;
                    break;
                case ePivotTableDatePeriodFilterType.M6:
                    df.Type = eDynamicFilterType.M6;
                    break;
                case ePivotTableDatePeriodFilterType.M7:
                    df.Type = eDynamicFilterType.M7;
                    break;
                case ePivotTableDatePeriodFilterType.M8:
                    df.Type = eDynamicFilterType.M8;
                    break;
                case ePivotTableDatePeriodFilterType.M9:
                    df.Type = eDynamicFilterType.M9;
                    break;
                case ePivotTableDatePeriodFilterType.M10:
                    df.Type = eDynamicFilterType.M10;
                    break;
                case ePivotTableDatePeriodFilterType.M11:
                    df.Type = eDynamicFilterType.M11;
                    break;
                case ePivotTableDatePeriodFilterType.M12:
                    df.Type = eDynamicFilterType.M12;
                    break;
                case ePivotTableDatePeriodFilterType.NextMonth:
                    df.Type = eDynamicFilterType.NextMonth;
                    break;
                case ePivotTableDatePeriodFilterType.NextQuarter:
                    df.Type = eDynamicFilterType.NextQuarter;
                    break;
                case ePivotTableDatePeriodFilterType.NextWeek:
                    df.Type = eDynamicFilterType.NextWeek;
                    break;
                case ePivotTableDatePeriodFilterType.NextYear:
                    df.Type = eDynamicFilterType.NextYear;
                    break;
                case ePivotTableDatePeriodFilterType.Q1:
                    df.Type = eDynamicFilterType.Q1;
                    break;
                case ePivotTableDatePeriodFilterType.Q2:
                    df.Type = eDynamicFilterType.Q2;
                    break;
                case ePivotTableDatePeriodFilterType.Q3:
                    df.Type = eDynamicFilterType.Q3;
                    break;
                case ePivotTableDatePeriodFilterType.Q4:
                    df.Type = eDynamicFilterType.Q4;
                    break;
                case ePivotTableDatePeriodFilterType.ThisMonth:
                    df.Type = eDynamicFilterType.ThisMonth;
                    break;
                case ePivotTableDatePeriodFilterType.ThisQuarter:
                    df.Type = eDynamicFilterType.ThisQuarter;
                    break;
                case ePivotTableDatePeriodFilterType.ThisWeek:
                    df.Type = eDynamicFilterType.ThisWeek;
                    break;
                case ePivotTableDatePeriodFilterType.ThisYear:
                    df.Type = eDynamicFilterType.ThisYear;
                    break;
                case ePivotTableDatePeriodFilterType.Yesterday:
                    df.Type = eDynamicFilterType.Yesterday;
                    break;
                case ePivotTableDatePeriodFilterType.Today:
                    df.Type = eDynamicFilterType.Today;
                    break;
                case ePivotTableDatePeriodFilterType.Tomorrow:
                    df.Type = eDynamicFilterType.Tomorrow;
                    break;
                case ePivotTableDatePeriodFilterType.YearToDate:
                    df.Type = eDynamicFilterType.YearToDate;
                    break;
                default:
                    throw new Exception($"Unsupported Pivottable filter type {type}");
            }

            _filter = df;
        }

        internal void CreateTop10Filter(ePivotTableTop10FilterType type, bool isTop, double value)
        {
            _filterColumnNode.InnerXml = "<top10 />";
            var tf = new ExcelTop10FilterColumn(NameSpaceManager, _filterColumnNode);

            tf.Percent = (type == ePivotTableTop10FilterType.Percent);
            tf.Top = isTop;
            tf.Value = value;
            tf.FilterValue = value;

            _filter = tf;
        }

        internal void CreateCaptionCustomFilter(ePivotTableCaptionFilterType type)
        {
            _filterColumnNode.InnerXml = "<customFilters/>";
            var cf = new ExcelCustomFilterColumn(NameSpaceManager, _filterColumnNode);

            eFilterOperator t;
            var v = StringValue1;
            switch (type)
            {
                case ePivotTableCaptionFilterType.CaptionNotBeginsWith:
                case ePivotTableCaptionFilterType.CaptionNotContains:
                case ePivotTableCaptionFilterType.CaptionNotEndsWith:
                case ePivotTableCaptionFilterType.CaptionNotEqual:
                    t = eFilterOperator.NotEqual;
                    break;
                case ePivotTableCaptionFilterType.CaptionGreaterThan:
                    t = eFilterOperator.GreaterThan;
                    break;
                case ePivotTableCaptionFilterType.CaptionGreaterThanOrEqual:
                case ePivotTableCaptionFilterType.CaptionBetween:
                    t = eFilterOperator.GreaterThanOrEqual;
                    break;
                case ePivotTableCaptionFilterType.CaptionLessThan:
                case ePivotTableCaptionFilterType.CaptionNotBetween:
                    t = eFilterOperator.LessThan;
                    break;
                case ePivotTableCaptionFilterType.CaptionLessThanOrEqual:
                    t = eFilterOperator.LessThanOrEqual;
                    break;
                default:
                    t = eFilterOperator.Equal;
                    break;
            }
            switch (type)
            {
                case ePivotTableCaptionFilterType.CaptionBeginsWith:
                case ePivotTableCaptionFilterType.CaptionNotBeginsWith:
                    v += "*";
                    break;
                case ePivotTableCaptionFilterType.CaptionContains:
                case ePivotTableCaptionFilterType.CaptionNotContains:
                    v = $"*{v}*";
                    break;
                case ePivotTableCaptionFilterType.CaptionEndsWith:
                case ePivotTableCaptionFilterType.CaptionNotEndsWith:
                    v = $"*{v}";
                    break;
            }
            var item1 = new ExcelFilterCustomItem(v, t);
            cf.Filters.Add(item1);

            if (type == ePivotTableCaptionFilterType.CaptionBetween)
            {
                cf.And = true;
                cf.Filters.Add(new ExcelFilterCustomItem(StringValue2, eFilterOperator.LessThanOrEqual));
            }
            else if (type == ePivotTableCaptionFilterType.CaptionNotBetween)
            {
                cf.And = false;
                cf.Filters.Add(new ExcelFilterCustomItem(StringValue2, eFilterOperator.GreaterThan));
            }

            _filter = cf;
        }
        internal void CreateValueCustomFilter(ePivotTableValueFilterType type)
        {
            _filterColumnNode.InnerXml = "<customFilters/>";
            var cf = new ExcelCustomFilterColumn(NameSpaceManager, _filterColumnNode);

            eFilterOperator t;
            string v1 = GetFilterValueAsString(Value1);
            switch (type)
            {
                case ePivotTableValueFilterType.ValueNotEqual:
                    t = eFilterOperator.NotEqual;
                    break;
                case ePivotTableValueFilterType.ValueGreaterThan:
                    t = eFilterOperator.GreaterThan;
                    break;
                case ePivotTableValueFilterType.ValueGreaterThanOrEqual:
                case ePivotTableValueFilterType.ValueBetween:
                    t = eFilterOperator.GreaterThanOrEqual;
                    break;
                case ePivotTableValueFilterType.ValueLessThan:
                    t = eFilterOperator.LessThan;
                    break;
                case ePivotTableValueFilterType.ValueLessThanOrEqual:
                case ePivotTableValueFilterType.ValueNotBetween:
                    t = eFilterOperator.LessThanOrEqual;
                    break;
                default:
                    t = eFilterOperator.Equal;
                    break;
            }

            var item1 = new ExcelFilterCustomItem(v1, t);
            cf.Filters.Add(item1);

            if (type == ePivotTableValueFilterType.ValueBetween)
            {
                cf.And = true;
                cf.Filters.Add(new ExcelFilterCustomItem(GetFilterValueAsString(Value2), eFilterOperator.LessThanOrEqual));
            }
            else if (type == ePivotTableValueFilterType.ValueNotBetween)
            {
                cf.And = false;
                cf.Filters.Add(new ExcelFilterCustomItem(GetFilterValueAsString(Value2), eFilterOperator.GreaterThan));
            }
            _filter = cf;
        }

        private string GetFilterValueAsString(object v)
        {
            if (ConvertUtil.IsNumericOrDate(v))
            {
                return ConvertUtil.GetValueDouble(v).ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                return v.ToString();
            }
        }
        internal void CreateValueFilter()
        {
            _filterColumnNode.InnerXml = "<filters/>";
            var f = new ExcelValueFilterColumn(NameSpaceManager, _filterColumnNode);
            f.Filters.Add(StringValue1);
            _filter = f;
        }

        internal bool MatchesLabel(ExcelPivotTable pivotTable, PivotTableCacheRecords recs, int index)
        {
            var t = (int)Type;
            var recordIndex = (int)recs.CacheItems[Fld][index];

            if (t < 100)     //Caption(string)
            {
                return MatchCaptions(pivotTable, recordIndex);
            }
            else if (t < 300) //Date
            {
                return MatchDate(pivotTable, recordIndex);
            }
            return false;
        }
		/// <summary>
		/// Handle caption(String) filters and the unknown filter type. This is filter enum values below 100.
		/// </summary>
		/// <param name="pivotTable"></param>
		/// <param name="index"></param>
		/// <returns></returns>
		private bool MatchCaptions(ExcelPivotTable pivotTable, int index)
    {
        var value = pivotTable.Fields[Fld].Cache.SharedItems[index].ToString();
        switch (Type)
        {
            //Caption filters (String)
            case ePivotTableFilterType.CaptionEqual:
                return value.Equals(StringValue1, StringComparison.InvariantCultureIgnoreCase);
            case ePivotTableFilterType.CaptionNotEqual:
                return !value.Equals(StringValue1, StringComparison.InvariantCultureIgnoreCase);
            case ePivotTableFilterType.CaptionBeginsWith:
                return value.StartsWith(StringValue1, StringComparison.InvariantCultureIgnoreCase);
            case ePivotTableFilterType.CaptionNotBeginsWith:
                return !value.StartsWith(StringValue1, StringComparison.InvariantCultureIgnoreCase);
            case ePivotTableFilterType.CaptionEndsWith:
                return value.EndsWith(StringValue1, StringComparison.InvariantCultureIgnoreCase);
            case ePivotTableFilterType.CaptionNotEndsWith:
                return !value.EndsWith(StringValue1, StringComparison.InvariantCultureIgnoreCase);
            case ePivotTableFilterType.CaptionGreaterThan:
                return string.Compare(value, StringValue1, true) > 0;
            case ePivotTableFilterType.CaptionGreaterThanOrEqual:
                return string.Compare(value, StringValue1, true) >= 0;
            case ePivotTableFilterType.CaptionLessThan:
                return string.Compare(value, StringValue1, true) < 0;
            case ePivotTableFilterType.CaptionLessThanOrEqual:
                return string.Compare(value, StringValue1, true) <= 0;
            case ePivotTableFilterType.CaptionContains:
                return value.IndexOf(StringValue1, StringComparison.InvariantCultureIgnoreCase) >= 0;
            case ePivotTableFilterType.CaptionNotContains:
                return value.IndexOf(StringValue1, StringComparison.InvariantCultureIgnoreCase) < 0;
            case ePivotTableFilterType.CaptionBetween:
                return string.Compare(value, StringValue1, true) >= 0 && string.Compare(value, StringValue2, true) <= 0;
            case ePivotTableFilterType.CaptionNotBetween:
                return !(string.Compare(value, StringValue1, true) >= 0 && string.Compare(value, StringValue2, true) <= 0);
            case ePivotTableFilterType.Unknown:
            default:
                return false;
        }
    }
        /// <summary>
        /// Handle date filters. This is filter enum values below 100.
        /// </summary>
        /// <param name="pivotTable"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        private bool MatchDate(ExcelPivotTable pivotTable, int index)
        {
            var date = ConvertUtil.GetValueDate(pivotTable.Fields[Fld].Cache.SharedItems[index]);
            var value1Date = ConvertUtil.GetValueDate(Value1);
            if (date.HasValue && value1Date.HasValue)
            {
                var compareDate = value1Date.Value.Date;
                switch (Type)
                {
                    case ePivotTableFilterType.DateEqual:
                        return date.Equals(compareDate);
                    case ePivotTableFilterType.DateNotEqual:
                        return !date.Equals(compareDate);
                    case ePivotTableFilterType.DateNewerThan:
                        return date > compareDate;
                    case ePivotTableFilterType.DateNewerThanOrEqual:
                        return date >= compareDate;
                    case ePivotTableFilterType.DateOlderThan:
                        return date < compareDate;
                    case ePivotTableFilterType.DateOlderThanOrEqual:
                        return date <= compareDate;
                    case ePivotTableFilterType.DateBetween:
                        return date >= compareDate && date <= (ConvertUtil.GetValueDate(Value2) ?? DateTime.MaxValue);
                    case ePivotTableFilterType.DateNotBetween:
                        return !(date >= compareDate && date <= (ConvertUtil.GetValueDate(Value2) ?? DateTime.MaxValue));
                    case ePivotTableFilterType.YearToDate:
                        return date.Value.Year == DateTime.Today.Year;
                    case ePivotTableFilterType.LastYear:
                        return date.Value.Year == DateTime.Today.Year - 1;
                    case ePivotTableFilterType.LastQuarter:
                        DateTimeUtil.GetQuarterDates(DateTime.Today.AddMonths(-3), out DateTime startDate, out DateTime endDate);
                        return date.Value >= startDate && date.Value <= endDate;
                    case ePivotTableFilterType.LastMonth:
                        var pm = DateTime.Today.AddMonths(-1);
                        return date.Value.Year == pm.Year && date.Value.Month == pm.Month;
                    case ePivotTableFilterType.LastWeek:
                        DateTimeUtil.GetWeekDates(DateTime.Today.AddDays(-7), out startDate, out endDate);
                        return date.Value >= startDate && date.Value <= endDate;
                    case ePivotTableFilterType.ThisYear:
                        return date.Value.Year == DateTime.Today.Year;
                    case ePivotTableFilterType.ThisQuarter:
                        DateTimeUtil.GetQuarterDates(DateTime.Today, out startDate, out endDate);
                        return date.Value >= startDate && date.Value <= endDate;
                    case ePivotTableFilterType.ThisMonth:
                        return date.Value.Year == DateTime.Today.Year && date.Value.Month == DateTime.Today.Month;
                    case ePivotTableFilterType.ThisWeek:
                        DateTimeUtil.GetWeekDates(DateTime.Today, out startDate, out endDate);
                        return date.Value >= startDate && date.Value <= endDate;
                    case ePivotTableFilterType.NextYear:
                        return date.Value.Year == DateTime.Today.Year + 1;
                    case ePivotTableFilterType.NextQuarter:
                        DateTimeUtil.GetQuarterDates(DateTime.Today.AddMonths(3), out startDate, out endDate);
                        return date.Value >= startDate && date.Value <= endDate;
                    case ePivotTableFilterType.NextMonth:
                        var nm = DateTime.Today.AddMonths(1);
                        return date.Value.Year == nm.Year && date.Value.Month == nm.Month;
                    case ePivotTableFilterType.NextWeek:
                        DateTimeUtil.GetWeekDates(DateTime.Today.AddDays(7), out startDate, out endDate);
                        return date.Value >= startDate && date.Value <= endDate;
                    case ePivotTableFilterType.M1:
                        return date.Value.Month == 1;
                    case ePivotTableFilterType.M2:
                        return date.Value.Month == 2;
                    case ePivotTableFilterType.M3:
                        return date.Value.Month == 3;
                    case ePivotTableFilterType.M4:
                        return date.Value.Month == 4;
                    case ePivotTableFilterType.M5:
                        return date.Value.Month == 5;
                    case ePivotTableFilterType.M6:
                        return date.Value.Month == 6;
                    case ePivotTableFilterType.M7:
                        return date.Value.Month == 7;
                    case ePivotTableFilterType.M8:
                        return date.Value.Month == 8;
                    case ePivotTableFilterType.M9:
                        return date.Value.Month == 9;
                    case ePivotTableFilterType.M10:
                        return date.Value.Month == 10;
                    case ePivotTableFilterType.M11:
                        return date.Value.Month == 11;
                    case ePivotTableFilterType.M12:
                        return date.Value.Month == 12;
                    case ePivotTableFilterType.Q1:
                        return date.Value.Month >= 1 && date.Value.Month <= 3;
                    case ePivotTableFilterType.Q2:
                        return date.Value.Month >= 4 && date.Value.Month <= 6;
                    case ePivotTableFilterType.Q3:
                        return date.Value.Month >= 7 && date.Value.Month <= 9;
                    case ePivotTableFilterType.Q4:
                        return date.Value.Month >= 10 && date.Value.Month <= 12;
                    case ePivotTableFilterType.Yesterday:
                        return date.Value.Date == DateTime.Today.AddDays(-1);
                    case ePivotTableFilterType.Today:
                        return date.Value.Date == DateTime.Today;
                    case ePivotTableFilterType.Tomorrow:
                        return date.Value.Date == DateTime.Today.AddDays(1);
                    default:
                        throw new InvalidOperationException($"Unknown date filter type {Type}");
                }
            }
            return false;
        }
        internal bool MatchNumeric(object value)
        {
            var num = RoundingHelper.RoundToSignificantFig(ConvertUtil.GetValueDouble(value), 15); 
            var value1 = RoundingHelper.RoundToSignificantFig(ConvertUtil.GetValueDouble(Value1, false, true), 15);
            if (double.IsNaN(num) == false && double.IsNaN(value1) == false)
            {
                switch (Type)
                {
                    case ePivotTableFilterType.ValueEqual:
                        return num.Equals(value1);
                    case ePivotTableFilterType.ValueNotEqual:
                        return !num.Equals(value1);
                    case ePivotTableFilterType.ValueGreaterThan:
                        return num > value1;
                    case ePivotTableFilterType.ValueGreaterThanOrEqual:
                        return num >= value1;
                    case ePivotTableFilterType.ValueLessThan:
                        return num < value1;
                    case ePivotTableFilterType.ValueLessThanOrEqual:
                        return num <= value1;
                    case ePivotTableFilterType.ValueBetween:
                        var value2 = ConvertUtil.GetValueDouble(Value2, false, true);
                        return double.IsNaN(value2)==false && num >= value1 && num <= value2;
                    case ePivotTableFilterType.ValueNotBetween:
                        value2 = ConvertUtil.GetValueDouble(Value2, false, true);
                        return double.IsNaN(value2) == false && !(num >= value1 && num <= value2);
                    default:
                        throw new InvalidOperationException($"Unknown date filter type {Type}");
                }
            }
            return false;
        }
        /// <summary>
        /// The type of pivot filter
        /// </summary>
        public ePivotTableFilterType Type
        {
            get
            {
                return GetXmlNodeString("@type").ToEnum(ePivotTableFilterType.Unknown);
            }
            internal set
            {
                var s = value.ToEnumString();
                if (s.Length <= 3 && (s[0]=='m' || s[0] == 'q')) s = s.ToUpper();  //For M1 - M12 and Q1 - Q4
                SetXmlNodeString("@type", s);
            }
        }
        /// <summary>
        /// The evaluation order of the pivot filter
        /// </summary>
        public int EvalOrder
        {
            get
            {
                return GetXmlNodeInt("@evalOrder");
            }
            internal set
            {
                SetXmlNodeInt("@evalOrder", value);
            }
        }
        /// <summary>
        /// The index to the row/column field the filter is applied on 
        /// </summary>
        internal int Fld
        {
            get
            {
                return GetXmlNodeInt("@fld");
            }
            set
            {
                SetXmlNodeInt("@fld", value);
            }
        }
        /// <summary>
        /// The index to the data field a value field is evaluated on.
        /// </summary>
        internal int MeasureFldIndex
        {
            get
            {
                return GetXmlNodeInt("@iMeasureFld");
            }
            set
            {
                SetXmlNodeInt("@iMeasureFld", value);
            }
        }
        internal int MeasureHierIndex
        {
            get
            {
                return GetXmlNodeInt("@iMeasureHier");
            }
            set
            {
                SetXmlNodeInt("@iMeasureHier", value);
            }
        }
        internal int MemberPropertyFldIndex
        {
            get
            {
                return GetXmlNodeInt("@mpFld");
            }
            set
            {
                SetXmlNodeInt("@mpFld", value);
            }
        }
        /// <summary>
        /// The valueOrIndex 1 to compare the filter to
        /// </summary>
        public object Value1
        {
            get;
            set;
        }

        /// <summary>
        /// The valueOrIndex 2 to compare the filter to
        /// </summary>
        public object Value2
        {
            get;
            set;
        }
        /// <summary>
        /// The string valueOrIndex 1 used by caption filters.
        /// </summary>
        internal string StringValue1
        {
            get
            {
                return GetXmlNodeString("@stringValue1");
            }
            set
            {
                SetXmlNodeString("@stringValue1", value, true);
            }
        }
        /// <summary>
        /// The string valueOrIndex 2 used by caption filters.
        /// </summary>
        internal string StringValue2
        {
            get
            {
                return GetXmlNodeString("@stringValue2");
            }
            set
            {
                SetXmlNodeString("@stringValue2", value, true);
            }
        }
        ExcelFilterColumn _filter = null;
        internal ExcelFilterColumn Filter
        {
            get
            {
                if (_filter == null)
                {
                    var filterNode = GetNode("d:autoFilter/d:filterColumn");
                    if (filterNode != null)
                    {
                        switch (filterNode.LocalName)
                        {
                            case "customFilters":
                                _filter = new ExcelCustomFilterColumn(NameSpaceManager, filterNode);
                                break;
                            case "top10":
                                _filter = new ExcelTop10FilterColumn(NameSpaceManager, filterNode);
                                break;
                            case "filters":
                                _filter = new ExcelValueFilterColumn(NameSpaceManager, filterNode);
                                break;
                            case "dynamicFilter":
                                _filter = new ExcelDynamicFilterColumn(NameSpaceManager, filterNode);
                                break;
                            case "colorFilter":
                                _filter = new ExcelColorFilterColumn(NameSpaceManager, filterNode);
                                break;
                            case "iconFilter":
                                _filter = new ExcelIconFilterColumn(NameSpaceManager, filterNode);
                                break;
                            default:
                                _filter = null;
                                break;
                        }
                    }
                    else
                    {
                        throw new Exception("Invalid xml in pivot table. Missing Filter column");
                    }
                }
                return _filter;
            }
            set
            {
                _filter = value;
            }
        }
    }
}
