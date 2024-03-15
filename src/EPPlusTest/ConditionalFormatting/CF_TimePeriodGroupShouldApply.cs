using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;

namespace EPPlusTest.ConditionalFormatting
{
    [TestClass]
    public class CF_TimePeriodGroupShouldApply : ShouldApplyTestBase
    {
        [TestMethod]
        public void CF_Last7DaysShouldApply()
        {
            var ws = CreatePackageSheet("Last7Days");
            var cf = ws.ConditionalFormatting.AddLast7Days(new ExcelAddress("A1:A5"));

            var today = DateTime.Today;

            AssertConditionalFormat(ws, today, (ExcelConditionalFormattingLast7Days)cf);
        }

        [TestMethod]
        public void CF_LastMonthShouldApply()
        {
            var ws = CreatePackageSheet("LastMonth");
            var cf = ws.ConditionalFormatting.AddLastMonth(new ExcelAddress("A1:A5"));

            var lastMonth = DateTime.Today.AddMonths(-1);

            AssertConditionalFormat(ws, lastMonth, (ExcelConditionalFormattingLastMonth)cf);
        }

        [TestMethod]
        public void CF_NextMonthShouldApply()
        {
            var ws = CreatePackageSheet("NextMonth");
            var cf = ws.ConditionalFormatting.AddNextMonth(new ExcelAddress("A1:A5"));

            var nextMonth = DateTime.Today.AddMonths(+1);

            AssertConditionalFormat(ws, nextMonth, (ExcelConditionalFormattingNextMonth)cf);
        }

        [TestMethod]
        public void CF_ThisMonthShouldApply()
        {
            var ws = CreatePackageSheet("ThisMonth");
            var cf = ws.ConditionalFormatting.AddThisMonth(new ExcelAddress("A1:A5"));

            var thisMonth = DateTime.Today;

            AssertConditionalFormat(ws, thisMonth, (ExcelConditionalFormattingThisMonth)cf);
        }

        [TestMethod]
        public void CF_LastWeekShouldApply()
        {
            var ws = CreatePackageSheet("LastWeek");
            var cf = ws.ConditionalFormatting.AddLastWeek(new ExcelAddress("A1:A5"));

            var lastWeek = DateTime.Today.AddDays(-7);

            AssertConditionalFormat(ws, lastWeek, (ExcelConditionalFormattingLastWeek)cf);
        }

        [TestMethod]
        public void CF_NextWeekShouldApply()
        {
            var ws = CreatePackageSheet("NextWeek");
            var cf = ws.ConditionalFormatting.AddNextWeek(new ExcelAddress("A1:A5"));

            var nextWeek = DateTime.Today.AddDays(+7);

            AssertConditionalFormat(ws, nextWeek, (ExcelConditionalFormattingNextWeek)cf);
        }

        [TestMethod]
        public void CF_ThisWeekShouldApply()
        {
            var ws = CreatePackageSheet("ThisWeek");
            var cf = ws.ConditionalFormatting.AddThisWeek(new ExcelAddress("A1:A5"));

            var thisWeek = DateTime.Today;

            AssertConditionalFormat(ws, thisWeek, (ExcelConditionalFormattingThisWeek)cf);
        }

        [TestMethod]
        public void CF_TodayShouldApply()
        {
            var ws = CreatePackageSheet("Today");
            var cf = ws.ConditionalFormatting.AddToday(new ExcelAddress("A1:A5"));

            var today = DateTime.Today;

            AssertConditionalFormat(ws, today, (ExcelConditionalFormattingToday)cf);
        }

        [TestMethod]
        public void CF_TomorrowShouldApply()
        {
            var ws = CreatePackageSheet("Tomorrow");
            var cf = ws.ConditionalFormatting.AddTomorrow(new ExcelAddress("A1:A5"));

            var yesterday = DateTime.Today.AddDays(+1);

            AssertConditionalFormat(ws, yesterday, (ExcelConditionalFormattingTomorrow)cf);
        }

        [TestMethod]
        public void CF_YesterdayShouldApply()
        {
            var ws = CreatePackageSheet("Yesterday");
            var cf = ws.ConditionalFormatting.AddYesterday(new ExcelAddress("A1:A5"));

            var yesterday = DateTime.Today.AddDays(-1);

            AssertConditionalFormat(ws, yesterday, (ExcelConditionalFormattingYesterday)cf);
        }
    }
}
