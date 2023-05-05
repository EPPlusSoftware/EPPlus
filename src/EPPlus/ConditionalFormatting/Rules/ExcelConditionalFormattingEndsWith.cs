using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingEndsWith : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingEndsWith
    {
        /****************************************************************************************/

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingEndsWith(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(
                eExcelConditionalFormattingRuleType.EndsWith,
                address,
                priority,
                worksheet
                )
        {
            UpdateFormula();
            Operator = eExcelConditionalFormattingOperatorType.EndsWith;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingEndsWith(
          ExcelAddress address,
          ExcelWorksheet worksheet,
          XmlReader xr)
          : base(
                eExcelConditionalFormattingRuleType.EndsWith,
                address,
                worksheet,
                xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.EndsWith;
        }

        internal ExcelConditionalFormattingEndsWith(ExcelConditionalFormattingEndsWith copy) : base(copy)
        {
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingEndsWith(this);
        }

        /// <summary>
        /// The text to search in the end of the cell
        /// </summary>
        public string ContainText
        {
            get
            {
                return Text;
            }
            set
            {
                Text = value;

                Formula = string.Format(
                  "RIGHT({0},LEN(\"{1}\"))=\"{1}\"",
                  Address.Start.Address,
                  value.Replace("\"", "\"\""));
            }
        }

        public override ExcelAddress Address 
        { 
            get { return base.Address; } 
            set { base.Address = value; UpdateFormula(); } 
        }

        void UpdateFormula()
        {
            Formula = string.Format(
              "RIGHT({0},LEN(\"{1}\"))=\"{1}\"",
              Address.Start.Address,
              Text);
        }

        #endregion Constructors

        /****************************************************************************************/
    }
}
