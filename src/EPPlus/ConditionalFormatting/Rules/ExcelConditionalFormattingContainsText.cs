
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingContainsText : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingContainsText
    {
        public ExcelConditionalFormattingContainsText(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.ContainsText, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.ContainsText;
            Text = string.Empty;
        }

        public ExcelConditionalFormattingContainsText(
          ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(eExcelConditionalFormattingRuleType.ContainsText, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.ContainsText;
        }

        ExcelConditionalFormattingContainsText(ExcelConditionalFormattingContainsText copy) :base(copy)
        {
            //Text = copy.Text;
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingContainsText(this);
        }

        internal override bool IsExtLst {
            get
            {
                if (Formula2 != null)
                {
                    return true;
                }

                return base.IsExtLst;
            }
        }

        public string Text
        {
            get
            {
                return _text;
            }
            set
            {
                _text = value;
                Formula2 = null;

                //TODO: Error check/Throw when formula does not follow this format and is a ContainsText.
                base.Formula = string.Format(
                  "NOT(ISERROR(SEARCH(\"{1}\",{0})))",
                  Address.Start.Address,
                  value.Replace("\"", "\"\""));
            }
        }

        //get Returns Formula2 and set sets both Formula and Formula2
        //Property name is Formula for Interface ease of use.
        //It is recommended to use the interface over cast when possible.
        public override string Formula
        {
            get
            {
                //We use Formula2 to store user input.
                //This because Formula has to be in a specific format for this class.
                return Formula2;
            }
            set
            {
                _text = null;
                Formula2 = value;

                //Set Formula to the required format with the Formula2 user input.
                base.Formula = string.Format(
                  "NOT(ISERROR(SEARCH({0},{1})))",
                  Formula2, Address.Start.Address);
            }
        }

        void UpdateFormula()
        {
            if (_text != null)
            {
                base.Formula = string.Format(
                  "NOT(ISERROR(SEARCH(\"{1}\",{0})))",
                  Address.Start.Address,
                  _text);
            }
            else if(Formula2 != null) 
            {
                Formula = Formula2;
            }
        }

        public override ExcelAddress Address
        {
            get { return base.Address; }
            set { base.Address = value; UpdateFormula(); }
        }
    }
}
