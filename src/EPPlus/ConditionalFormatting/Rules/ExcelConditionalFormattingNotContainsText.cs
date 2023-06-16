
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingNotContainsText : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingNotContainsText
    {
        public ExcelConditionalFormattingNotContainsText(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.NotContainsText, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.NotContains;
            ContainText = string.Empty;
        }

        public ExcelConditionalFormattingNotContainsText(
          ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(eExcelConditionalFormattingRuleType.ContainsText, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.NotContains;
            string test = Formula.Substring(16, 2);
            if (test == "\"")
            {
                Text = Formula.GetSubstringStoppingAtSymbol("ISERROR(SEARCH(\"".Length);
            }
            else
            {
                //TODO: Will this create problems in different cultures with different seperators?
                _formulaReference = Formula.GetSubstringStoppingAtSymbol("ISERROR(SEARCH(".Length,",");
                Formula2 = _formulaReference;
            }
        }

        ExcelConditionalFormattingNotContainsText(ExcelConditionalFormattingNotContainsText copy) :base(copy)
        {
            ContainText = copy.ContainText;
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingNotContainsText(this);
        }

        public string ContainText
        {
            get
            {
                return Text;
            }
            set
            {
                Text = value;
                _formulaReference = null;
                Formula2 = null;

                //TODO: Error check/Throw when formula does not follow this format and is a ContainsText.
                Formula = string.Format(
                  "ISERROR(SEARCH(\"{1}\",{0}))",
                  Address.Start.Address,
                  value.Replace("\"", "\"\""));
            }
        }

        string _formulaReference = null;

        public string FormulaReference
        {
            get
            {
                return _formulaReference;
            }
            set
            {
                Text = null;
                _formulaReference = value;
                Formula2 = value;

                Formula = string.Format(
                    "ISERROR(SEARCH({1},{0}))",
                    Address.Start.Address,
                    value.Replace("\"", "\"\""));
            }
        }

        void UpdateFormula()
        {
            if (Text != null)
            {
                Formula = string.Format(
                    "ISERROR(SEARCH(\"{1}\",{0}))",
                    Address.Start.Address,
                    Text);
            }
            else if (_formulaReference != null)
            {
                Formula = string.Format(
                    "ISERROR(SEARCH({1},{0}))",
                    Address.Start.Address,
                    _formulaReference.Replace("\"", "\"\""));
            }
        }

        internal override bool IsExtLst
        {
            get
            {
                if (_formulaReference != null)
                {
                    return true;
                }

                return base.IsExtLst;
            }
        }

        public override ExcelAddress Address
        {
            get { return base.Address; }
            set { base.Address = value; UpdateFormula(); }
        }
    }
}
