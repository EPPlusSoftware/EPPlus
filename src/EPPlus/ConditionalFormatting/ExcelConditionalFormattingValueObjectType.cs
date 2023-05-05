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
	/// Functions related to the <see cref="ExcelConditionalFormattingColorScaleValueOLD"/>
	/// </summary>
	internal static class ExcelConditionalFormattingValueObjectType
	{
		/// <summary>
		/// Get the sequencial order of a cfvo/color by its position.
		/// </summary>
		/// <param name="position"></param>
        /// <param name="ruleType"></param>
		/// <returns>1, 2 or 3</returns>
		internal static int GetOrderByPosition(
			eExcelConditionalFormattingValueObjectPosition position,
			eExcelConditionalFormattingRuleType ruleType)
		{
			switch (position)
			{
				case eExcelConditionalFormattingValueObjectPosition.Low:
					return 1;

				case eExcelConditionalFormattingValueObjectPosition.Middle:
					return 2;

				case eExcelConditionalFormattingValueObjectPosition.High:
					// Check if the rule type is TwoColorScale.
					if (ruleType == eExcelConditionalFormattingRuleType.TwoColorScale)
					{
						// There are only "Low" and "High". So "High" is the second
						return 2;
					}

					// There are "Low", "Middle" and "High". So "High" is the third
					return 3;
			}

			return 0;
		}

		/// <summary>
		/// Get the CFVO type by its @type attribute
		/// </summary>
		/// <param name="attribute"></param>
		/// <returns></returns>
		public static eExcelConditionalFormattingValueObjectType GetTypeByAttrbiute(
			string attribute)
		{
			switch (attribute)
			{
				case ExcelConditionalFormattingConstants.CfvoType.Min:
					return eExcelConditionalFormattingValueObjectType.Min;

        case ExcelConditionalFormattingConstants.CfvoType.Max:
					return eExcelConditionalFormattingValueObjectType.Max;

        case ExcelConditionalFormattingConstants.CfvoType.Num:
					return eExcelConditionalFormattingValueObjectType.Num;

        case ExcelConditionalFormattingConstants.CfvoType.Formula:
					return eExcelConditionalFormattingValueObjectType.Formula;

        case ExcelConditionalFormattingConstants.CfvoType.Percent:
					return eExcelConditionalFormattingValueObjectType.Percent;

        case ExcelConditionalFormattingConstants.CfvoType.Percentile:
					return eExcelConditionalFormattingValueObjectType.Percentile;
			}

			throw new Exception(
        ExcelConditionalFormattingConstants.Errors.UnexistentCfvoTypeAttribute);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="position"></param>
		///<param name="ruleType"></param>
        /// <param name="topNode"></param>
		/// <param name="nameSpaceManager"></param>
		/// <returns></returns>
		public static XmlNode GetCfvoNodeByPosition(
			eExcelConditionalFormattingValueObjectPosition position,
			eExcelConditionalFormattingRuleType ruleType,
			XmlNode topNode,
			XmlNamespaceManager nameSpaceManager)
		{
			// Get the corresponding <cfvo> node (by the position)
			var node = topNode.SelectSingleNode(
				string.Format(
					"{0}[position()={1}]",
				// {0}
					ExcelConditionalFormattingConstants.Paths.Cfvo,
				// {1}
					ExcelConditionalFormattingValueObjectType.GetOrderByPosition(position, ruleType)),
				nameSpaceManager);

			if (node == null)
			{
				throw new Exception(
          ExcelConditionalFormattingConstants.Errors.MissingCfvoNode);
			}

			return node;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="type"></param>
		/// <returns></returns>
		public static string GetAttributeByType(
			eExcelConditionalFormattingValueObjectType type)
		{
			switch (type)
			{
				case eExcelConditionalFormattingValueObjectType.Min:
					return ExcelConditionalFormattingConstants.CfvoType.Min;

				case eExcelConditionalFormattingValueObjectType.Max:
          return ExcelConditionalFormattingConstants.CfvoType.Max;

				case eExcelConditionalFormattingValueObjectType.Num:
          return ExcelConditionalFormattingConstants.CfvoType.Num;

				case eExcelConditionalFormattingValueObjectType.Formula:
          return ExcelConditionalFormattingConstants.CfvoType.Formula;

				case eExcelConditionalFormattingValueObjectType.Percent:
          return ExcelConditionalFormattingConstants.CfvoType.Percent;

				case eExcelConditionalFormattingValueObjectType.Percentile:
          return ExcelConditionalFormattingConstants.CfvoType.Percentile;
			}

			return string.Empty;
		}

		/// <summary>
		/// Get the cfvo (§18.3.1.11) node parent by the rule type. Can be any of the following:
		/// "colorScale" (§18.3.1.16); "dataBar" (§18.3.1.28); "iconSet" (§18.3.1.49)
		/// </summary>
		/// <param name="ruleType"></param>
		/// <returns></returns>
		public static string GetParentPathByRuleType(
			eExcelConditionalFormattingRuleType ruleType)
		{
			switch (ruleType)
			{
				case eExcelConditionalFormattingRuleType.TwoColorScale:
				case eExcelConditionalFormattingRuleType.ThreeColorScale:
					return ExcelConditionalFormattingConstants.Paths.ColorScale;

				case eExcelConditionalFormattingRuleType.ThreeIconSet:
                case eExcelConditionalFormattingRuleType.FourIconSet:
                case eExcelConditionalFormattingRuleType.FiveIconSet:
					        return ExcelConditionalFormattingConstants.Paths.IconSet;

                case eExcelConditionalFormattingRuleType.DataBar:
                  return ExcelConditionalFormattingConstants.Paths.DataBar;
              }

			return string.Empty;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="nodeType"></param>
		/// <returns></returns>
		public static string GetNodePathByNodeType(
			eExcelConditionalFormattingValueObjectNodeType nodeType)
		{
			switch(nodeType)
			{
				case eExcelConditionalFormattingValueObjectNodeType.Cfvo:
					return ExcelConditionalFormattingConstants.Paths.Cfvo;

				case eExcelConditionalFormattingValueObjectNodeType.Color:
					return ExcelConditionalFormattingConstants.Paths.Color;
			}

			return string.Empty;
		}
	}
}