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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
	/// <summary>
  /// Functions related to the <see cref="ExcelConditionalFormattingTimePeriodType"/>
	/// </summary>
  internal static class ExcelConditionalFormattingTimePeriodType
	{
		/// <summary>
		/// 
		/// </summary>
		/// <param name="type"></param>
		/// <returns></returns>
		public static string GetAttributeByType(
			eExcelConditionalFormattingTimePeriodType type)
		{
			switch (type)
			{
        case eExcelConditionalFormattingTimePeriodType.Last7Days:
          return ExcelConditionalFormattingConstants.TimePeriods.Last7Days;

        case eExcelConditionalFormattingTimePeriodType.LastMonth:
          return ExcelConditionalFormattingConstants.TimePeriods.LastMonth;

        case eExcelConditionalFormattingTimePeriodType.LastWeek:
          return ExcelConditionalFormattingConstants.TimePeriods.LastWeek;

        case eExcelConditionalFormattingTimePeriodType.NextMonth:
          return ExcelConditionalFormattingConstants.TimePeriods.NextMonth;

        case eExcelConditionalFormattingTimePeriodType.NextWeek:
          return ExcelConditionalFormattingConstants.TimePeriods.NextWeek;

        case eExcelConditionalFormattingTimePeriodType.ThisMonth:
          return ExcelConditionalFormattingConstants.TimePeriods.ThisMonth;

        case eExcelConditionalFormattingTimePeriodType.ThisWeek:
          return ExcelConditionalFormattingConstants.TimePeriods.ThisWeek;

        case eExcelConditionalFormattingTimePeriodType.Today:
          return ExcelConditionalFormattingConstants.TimePeriods.Today;

        case eExcelConditionalFormattingTimePeriodType.Tomorrow:
          return ExcelConditionalFormattingConstants.TimePeriods.Tomorrow;

        case eExcelConditionalFormattingTimePeriodType.Yesterday:
          return ExcelConditionalFormattingConstants.TimePeriods.Yesterday;
			}

			return string.Empty;
		}

    /// <summary>
    /// 
    /// </summary>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static eExcelConditionalFormattingTimePeriodType GetTypeByAttribute(
      string attribute)
    {
      switch (attribute)
      {
        case ExcelConditionalFormattingConstants.TimePeriods.Last7Days:
          return eExcelConditionalFormattingTimePeriodType.Last7Days;

        case ExcelConditionalFormattingConstants.TimePeriods.LastMonth:
          return eExcelConditionalFormattingTimePeriodType.LastMonth;

        case ExcelConditionalFormattingConstants.TimePeriods.LastWeek:
          return eExcelConditionalFormattingTimePeriodType.LastWeek;

        case ExcelConditionalFormattingConstants.TimePeriods.NextMonth:
          return eExcelConditionalFormattingTimePeriodType.NextMonth;

        case ExcelConditionalFormattingConstants.TimePeriods.NextWeek:
          return eExcelConditionalFormattingTimePeriodType.NextWeek;

        case ExcelConditionalFormattingConstants.TimePeriods.ThisMonth:
          return eExcelConditionalFormattingTimePeriodType.ThisMonth;

        case ExcelConditionalFormattingConstants.TimePeriods.ThisWeek:
          return eExcelConditionalFormattingTimePeriodType.ThisWeek;

        case ExcelConditionalFormattingConstants.TimePeriods.Today:
          return eExcelConditionalFormattingTimePeriodType.Today;

        case ExcelConditionalFormattingConstants.TimePeriods.Tomorrow:
          return eExcelConditionalFormattingTimePeriodType.Tomorrow;

        case ExcelConditionalFormattingConstants.TimePeriods.Yesterday:
          return eExcelConditionalFormattingTimePeriodType.Yesterday;
      }

      throw new Exception(
        ExcelConditionalFormattingConstants.Errors.UnexistentTimePeriodTypeAttribute);
    }
  }
}