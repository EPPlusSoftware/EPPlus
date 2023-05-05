using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingNotEqual : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingNotEqual
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingNotEqual(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.NotEqual, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.NotEqual;
            Formula = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingNotEqual(ExcelAddress address, ExcelWorksheet ws, XmlReader xr) 
            : base(eExcelConditionalFormattingRuleType.NotEqual, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.NotEqual;
        }

        internal ExcelConditionalFormattingNotEqual(ExcelConditionalFormattingNotEqual copy) : base(copy)
        {
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingNotEqual(this);
        }

        #endregion Constructors
    }
}
