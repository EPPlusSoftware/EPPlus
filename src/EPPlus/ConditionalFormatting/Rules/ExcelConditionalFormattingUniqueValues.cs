using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingUniqueValues : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingUniqueValues
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingUniqueValues(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.UniqueValues, address, priority, worksheet)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingUniqueValues(ExcelAddress address, ExcelWorksheet ws, XmlReader xr) 
            : base(eExcelConditionalFormattingRuleType.UniqueValues, address, ws, xr)
        {
        }

        internal ExcelConditionalFormattingUniqueValues(ExcelConditionalFormattingUniqueValues copy) : base(copy)
        {
            Rank = copy.Rank;
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingUniqueValues(this);
        }

        #endregion Constructors
    }
}
