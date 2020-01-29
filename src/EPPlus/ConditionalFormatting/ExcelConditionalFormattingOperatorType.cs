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
  /// Functions related to the <see cref="ExcelConditionalFormattingOperatorType"/>
	/// </summary>
  internal static class ExcelConditionalFormattingOperatorType
	{
		/// <summary>
		/// 
		/// </summary>
		/// <param name="type"></param>
		/// <returns></returns>
		internal static string GetAttributeByType(
			eExcelConditionalFormattingOperatorType type)
		{
			switch (type)
			{
        case eExcelConditionalFormattingOperatorType.BeginsWith:
          return ExcelConditionalFormattingConstants.Operators.BeginsWith;

        case eExcelConditionalFormattingOperatorType.Between:
          return ExcelConditionalFormattingConstants.Operators.Between;

        case eExcelConditionalFormattingOperatorType.ContainsText:
          return ExcelConditionalFormattingConstants.Operators.ContainsText;

        case eExcelConditionalFormattingOperatorType.EndsWith:
          return ExcelConditionalFormattingConstants.Operators.EndsWith;

        case eExcelConditionalFormattingOperatorType.Equal:
          return ExcelConditionalFormattingConstants.Operators.Equal;

        case eExcelConditionalFormattingOperatorType.GreaterThan:
          return ExcelConditionalFormattingConstants.Operators.GreaterThan;

        case eExcelConditionalFormattingOperatorType.GreaterThanOrEqual:
          return ExcelConditionalFormattingConstants.Operators.GreaterThanOrEqual;

        case eExcelConditionalFormattingOperatorType.LessThan:
          return ExcelConditionalFormattingConstants.Operators.LessThan;

        case eExcelConditionalFormattingOperatorType.LessThanOrEqual:
          return ExcelConditionalFormattingConstants.Operators.LessThanOrEqual;

        case eExcelConditionalFormattingOperatorType.NotBetween:
          return ExcelConditionalFormattingConstants.Operators.NotBetween;

        case eExcelConditionalFormattingOperatorType.NotContains:
          return ExcelConditionalFormattingConstants.Operators.NotContains;

        case eExcelConditionalFormattingOperatorType.NotEqual:
          return ExcelConditionalFormattingConstants.Operators.NotEqual;
			}

			return string.Empty;
		}

    /// <summary>
    /// 
    /// </summary>
    /// param name="attribute"
    /// <returns></returns>
    internal static eExcelConditionalFormattingOperatorType GetTypeByAttribute(
      string attribute)
    {
      switch (attribute)
      {
        case ExcelConditionalFormattingConstants.Operators.BeginsWith:
          return eExcelConditionalFormattingOperatorType.BeginsWith;

        case ExcelConditionalFormattingConstants.Operators.Between:
          return eExcelConditionalFormattingOperatorType.Between;

        case ExcelConditionalFormattingConstants.Operators.ContainsText:
          return eExcelConditionalFormattingOperatorType.ContainsText;

        case ExcelConditionalFormattingConstants.Operators.EndsWith:
          return eExcelConditionalFormattingOperatorType.EndsWith;

        case ExcelConditionalFormattingConstants.Operators.Equal:
          return eExcelConditionalFormattingOperatorType.Equal;

        case ExcelConditionalFormattingConstants.Operators.GreaterThan:
          return eExcelConditionalFormattingOperatorType.GreaterThan;

        case ExcelConditionalFormattingConstants.Operators.GreaterThanOrEqual:
          return eExcelConditionalFormattingOperatorType.GreaterThanOrEqual;

        case ExcelConditionalFormattingConstants.Operators.LessThan:
          return eExcelConditionalFormattingOperatorType.LessThan;

        case ExcelConditionalFormattingConstants.Operators.LessThanOrEqual:
          return eExcelConditionalFormattingOperatorType.LessThanOrEqual;

        case ExcelConditionalFormattingConstants.Operators.NotBetween:
          return eExcelConditionalFormattingOperatorType.NotBetween;

        case ExcelConditionalFormattingConstants.Operators.NotContains:
          return eExcelConditionalFormattingOperatorType.NotContains;

        case ExcelConditionalFormattingConstants.Operators.NotEqual:
          return eExcelConditionalFormattingOperatorType.NotEqual;
      }

      throw new Exception(
        ExcelConditionalFormattingConstants.Errors.UnexistentOperatorTypeAttribute);
    }
  }
}