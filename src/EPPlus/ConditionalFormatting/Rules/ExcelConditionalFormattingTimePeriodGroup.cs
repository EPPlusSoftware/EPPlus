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
using System.Globalization;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

namespace OfficeOpenXml.ConditionalFormatting
{
  /// <summary>
  /// ExcelConditionalFormattingTimePeriodGroup
  /// </summary>
  internal class ExcelConditionalFormattingTimePeriodGroup: ExcelConditionalFormattingRule,
    IExcelConditionalFormattingTimePeriodGroup
  {
    /****************************************************************************************/

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="priority"></param>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingTimePeriodGroup(
            eExcelConditionalFormattingRuleType type,
            ExcelAddress address,
            int priority,
            ExcelWorksheet worksheet)
            : base(type, address, priority, worksheet)
        {
        }

        internal ExcelConditionalFormattingTimePeriodGroup(
            eExcelConditionalFormattingRuleType type,
            ExcelAddress address,
            ExcelWorksheet ws,
            XmlReader xr)
            : base(type, address, ws, xr)
        {
        }

        internal ExcelConditionalFormattingTimePeriodGroup(ExcelConditionalFormattingTimePeriodGroup copy, ExcelWorksheet newWs = null) : base(copy, newWs)
        {
            TimePeriod = copy.TimePeriod;
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingTimePeriodGroup(this, newWs);
        }

        protected string _baseFormula = null;

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            var formAtAddress = string.Format(
            _baseFormula,
            address.Start.Address);

            var formResult = _ws.Workbook.FormulaParserManager.Parse(formAtAddress, address.FullAddress, false);
            if(ExcelErrorValue.Values.IsErrorValue(formResult))
            {
                return false;
            }
            var formattedResult = string.Format(formResult.ToString(), CultureInfo.InvariantCulture);

            return bool.Parse(formattedResult);
        }


        #endregion Constructors

        /****************************************************************************************/
    }
}