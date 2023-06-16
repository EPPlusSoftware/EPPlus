
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
            ContainText = string.Empty;
        }

        public ExcelConditionalFormattingContainsText(
          ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(eExcelConditionalFormattingRuleType.ContainsText, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.ContainsText;

            if (Formula2 != null)
            {
                _formulaReference = Formula2;
            }
            else if (Text != null)
            {
                Text = Formula.GetSubstringStoppingAtSymbol("NOT(ISERROR(SEARCH(\"".Length);
            }
        }

        ExcelConditionalFormattingContainsText(ExcelConditionalFormattingContainsText copy) :base(copy)
        {
            ContainText = copy.ContainText;
        }

        internal override bool IsExtLst {
            get
            {
                if (_formulaReference != null)
                {
                    return true;
                }

                return base.IsExtLst;
            }
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingContainsText(this);
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
                  "NOT(ISERROR(SEARCH(\"{1}\",{0})))",
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
                  "NOT(ISERROR(SEARCH({1},{0})))",
                  Address.Start.Address, value);
            }
        }

        void UpdateFormula()
        {
            if (Text != null)
            {
                Formula = string.Format(
                  "NOT(ISERROR(SEARCH(\"{1}\",{0})))",
                  Address.Start.Address,
                  Text);
            }
            else if(_formulaReference != null) 
            {
                Formula = string.Format(
                "NOT(ISERROR(SEARCH({1},{0})))",
                Address.Start.Address,
                _formulaReference);
            }
        }

        public override ExcelAddress Address
        {
            get { return base.Address; }
            set { base.Address = value; UpdateFormula(); }
        }
    }
}
