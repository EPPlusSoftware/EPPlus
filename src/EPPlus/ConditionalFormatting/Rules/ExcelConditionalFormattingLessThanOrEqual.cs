using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingLessThanOrEqual : ExcelConditionalFormattingRule, IExcelConditionalFormattingLessThanOrEqual
    {

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingLessThanOrEqual(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.LessThanOrEqual, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.LessThanOrEqual;
            Formula = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingLessThanOrEqual(ExcelAddress address, ExcelWorksheet ws, XmlReader xr) 
            : base(eExcelConditionalFormattingRuleType.LessThanOrEqual, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.LessThanOrEqual;
        }

        internal ExcelConditionalFormattingLessThanOrEqual(ExcelConditionalFormattingLessThanOrEqual copy) : base(copy)
        {
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingLessThanOrEqual(this);
        }

        #endregion Constructors
    }
}
