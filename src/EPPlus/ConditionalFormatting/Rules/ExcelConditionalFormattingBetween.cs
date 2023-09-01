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
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingBetween : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingBetween
    {
        /****************************************************************************************/

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingBetween(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(
                eExcelConditionalFormattingRuleType.Between,
                address,
                priority,
                worksheet
                )
        {
            Operator = eExcelConditionalFormattingOperatorType.Between;
            Formula = string.Empty;
            Formula2 = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingBetween(
          ExcelAddress address,
          ExcelWorksheet worksheet,
          XmlReader xr)
          : base(
                eExcelConditionalFormattingRuleType.Between,
                address,
                worksheet,
                xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.Between;
        }

        internal ExcelConditionalFormattingBetween(ExcelConditionalFormattingBetween copy, ExcelWorksheet newWs = null) : base(copy, newWs)
        {
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingBetween(this, newWs);
        }


        #endregion Constructors

        /****************************************************************************************/
    }
}
