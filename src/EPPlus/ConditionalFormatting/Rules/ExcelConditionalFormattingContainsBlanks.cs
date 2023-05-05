using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingContainsBlanks : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingContainsBlanks
    {
        /****************************************************************************************/

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingContainsBlanks(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(
                eExcelConditionalFormattingRuleType.ContainsBlanks,
                address,
                priority,
                worksheet
                )
        {
            Formula = string.Format(
              "LEN(TRIM({0}))=0",
              Address.Start.Address);
        }

        internal ExcelConditionalFormattingContainsBlanks(ExcelConditionalFormattingContainsBlanks copy) : base(copy)
        {
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingContainsBlanks(this);
        }


        void UpdateFormula()
        {
            Formula = string.Format(
              "LEN(TRIM({0}))=0",
              Address.Start.Address);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingContainsBlanks(
          ExcelAddress address,
          ExcelWorksheet worksheet,
          XmlReader xr)
          : base(
                eExcelConditionalFormattingRuleType.ContainsBlanks,
                address,
                worksheet,
                xr)
        {
        }

        public override ExcelAddress Address 
        { 
            get { return base.Address; } 
            set { base.Address = value; UpdateFormula(); } 
        }


        #endregion Constructors

        /****************************************************************************************/
    }
}
