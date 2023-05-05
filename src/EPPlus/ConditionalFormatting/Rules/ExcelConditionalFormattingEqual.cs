using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingEqual : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingEqual
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingEqual(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.Equal, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.Equal;
            Formula = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingEqual(ExcelAddress address, ExcelWorksheet ws, XmlReader xr) 
            : base(eExcelConditionalFormattingRuleType.Equal, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.Equal;
        }
        internal ExcelConditionalFormattingEqual(ExcelConditionalFormattingEqual copy) : base(copy)
        {
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingEqual(this);
        }


        #endregion Constructors
    }
}
