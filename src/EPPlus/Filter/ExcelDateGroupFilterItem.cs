/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// A date group for filters
    /// </summary>
    public class ExcelFilterDateGroupItem : ExcelFilterItem
    {
        /// <summary>
        /// Filter out the specified year
        /// </summary>
        /// <param name="year">The year</param>
        public ExcelFilterDateGroupItem(int year)
        {
            Grouping = eDateTimeGrouping.Year;
            Year = year;
            Validate();
        }

        /// <summary>
        /// Filter out the specified year and month
        /// </summary>
        /// <param name="year">The year</param>
        /// <param name="month">The month</param>
        public ExcelFilterDateGroupItem(int year, int month)
        {
            Grouping = eDateTimeGrouping.Month;
            Year = year;
            Month = month;
            Validate();
        }
        /// <summary>
        /// Filter out the specified year, month and day
        /// </summary>
        /// <param name="year">The year</param>
        /// <param name="month">The month</param>
        /// <param name="day">The day</param>
        public ExcelFilterDateGroupItem(int year, int month, int day)
        {
            Grouping = eDateTimeGrouping.Day;
            Year = year;
            Month = month;
            Day = day;
            Validate();
        }
        /// <summary>
        /// Filter out the specified year, month, day and hour
        /// </summary>
        /// <param name="year">The year</param>
        /// <param name="month">The month</param>
        /// <param name="day">The day</param>
        /// <param name="hour">The hour</param>
        public ExcelFilterDateGroupItem(int year, int month, int day, int hour)
        {
            Grouping = eDateTimeGrouping.Hour;
            Year = year;
            Month = month;
            Day = day;
            Hour = hour;
            Validate();
        }
        /// <summary>
        /// Filter out the specified year, month, day, hour and and minute
        /// </summary>
        /// <param name="year">The year</param>
        /// <param name="month">The month</param>
        /// <param name="day">The day</param>
        /// <param name="hour">The hour</param>
        /// <param name="minute">The minute</param>
        public ExcelFilterDateGroupItem(int year, int month, int day, int hour, int minute)
        {
            Grouping = eDateTimeGrouping.Minute;
            Year = year;
            Month = month;
            Day = day;
            Hour = hour;
            Minute = minute;
            Validate();
        }
        /// <summary>
        /// Filter out the specified year, month, day, hour and and minute
        /// </summary>
        /// <param name="year">The year</param>
        /// <param name="month">The month</param>
        /// <param name="day">The day</param>
        /// <param name="hour">The hour</param>
        /// <param name="minute">The minute</param>
        /// <param name="second">The second</param>
        public ExcelFilterDateGroupItem(int year, int month, int day, int hour, int minute, int second)
        {
            Grouping = eDateTimeGrouping.Second;
            Year = year;
            Month = month;
            Day = day;
            Hour = hour;
            Minute = minute;
            Second = second;
            Validate();
        }
        private void Validate()
        {
            if (Year < 0 && Year > 9999)
            {
                throw (new ArgumentException("Year out of range(0-9999)"));
            }

            if (Grouping == eDateTimeGrouping.Year) return;
            if (Month < 1 && Month > 12)
            {
                throw (new ArgumentException("Month out of range(1-12)"));
            }
            if (Grouping == eDateTimeGrouping.Month) return;

            if (Day < 1 && Day > 31)
            {
                throw (new ArgumentException("Month out of range(1-31)"));
            }
            if (Grouping == eDateTimeGrouping.Day) return;

            if (Hour < 0 && Hour > 23)
            {
                throw (new ArgumentException("Hour out of range(0-23)"));
            }
            if (Grouping == eDateTimeGrouping.Hour) return;

            if (Minute < 0 && Minute > 59)
            {
                throw (new ArgumentException("Minute out of range(0-59)"));
            }
            if (Grouping == eDateTimeGrouping.Minute) return;

            if (Second < 0 && Second > 59)
            {
                throw (new ArgumentException("Second out of range(0-59)"));
            }
        }

        internal void AddNode(XmlNode node)
        {
            var e = node.OwnerDocument.CreateElement("dateGroupItem", ExcelPackage.schemaMain);
            e.SetAttribute("dateTimeGrouping", Grouping.ToString().ToLower());
            e.SetAttribute("year", Year.ToString(CultureInfo.InvariantCulture));

            if (Month.HasValue)
            {
                e.SetAttribute("month", Month.Value.ToString(CultureInfo.InvariantCulture));
                if (Day.HasValue)
                {
                    e.SetAttribute("day", Day.Value.ToString(CultureInfo.InvariantCulture));
                    if (Hour.HasValue)
                    {
                        e.SetAttribute("hour", Hour.Value.ToString(CultureInfo.InvariantCulture));
                        if (Minute.HasValue)
                        {
                            e.SetAttribute("minute", Minute.Value.ToString(CultureInfo.InvariantCulture));
                            if (Second.HasValue)
                            {
                                e.SetAttribute("second", Second.Value.ToString(CultureInfo.InvariantCulture));
                            }
                        }
                    }
                }
            }

            node.AppendChild(e);
        }

        internal bool Match(DateTime value)
        {
            var match = value.Year == Year;

            if(match && Month.HasValue)
            {
                match = value.Month == Month;
                if(match && Day.HasValue)
                {
                    match = value.Day == Day;
                    if (match && Hour.HasValue)
                    {
                        match = value.Hour == Hour;
                        if (match && Minute.HasValue)
                        {
                            match = value.Minute == Minute;
                            if (match && Second.HasValue)
                            {
                                match = value.Second == Second;
                            }
                        }
                    }
                }
            }
            return match;
        }
        /// <summary>
        /// The grouping. Is set depending on the selected constructor
        /// </summary>
        public eDateTimeGrouping Grouping{ get; }
        /// <summary>
        /// Year to filter on
        /// </summary>
        public int Year { get; }
        /// <summary>
        /// Month to filter on
        /// </summary>
        public int? Month { get; }
        /// <summary>
        /// Day to filter on
        /// </summary>
        public int? Day { get; }
        /// <summary>
        /// Hour to filter on
        /// </summary>
        public int? Hour { get; }
        /// <summary>
        /// Minute to filter on
        /// </summary>
        public int? Minute { get;  }
        /// <summary>
        /// Second to filter on
        /// </summary>
        public int? Second { get;  }
    }
}