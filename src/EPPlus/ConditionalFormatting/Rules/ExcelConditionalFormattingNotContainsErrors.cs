using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingNotContainsErrors : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingNotContainsErrors
    {
        /****************************************************************************************/

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingNotContainsErrors(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(
                eExcelConditionalFormattingRuleType.NotContainsErrors,
                address,
                priority,
                worksheet
                )
        {
            UpdateFormula();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingNotContainsErrors(
          ExcelAddress address,
          ExcelWorksheet worksheet,
          XmlReader xr)
          : base(
                eExcelConditionalFormattingRuleType.NotContainsErrors,
                address,
                worksheet,
                xr)
        {
        }

        internal ExcelConditionalFormattingNotContainsErrors(ExcelConditionalFormattingNotContainsErrors copy) : base(copy)
        {
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingNotContainsErrors(this);
        }

        public override ExcelAddress Address 
        { 
            get { return base.Address; } 
            set { base.Address = value; UpdateFormula(); } 
        }

        void UpdateFormula()
        {
            Formula = string.Format(
              "NOT(ISERROR({0}))",
              Address.Start.Address);
        }

        #endregion Constructors

        /****************************************************************************************/
    }
}
