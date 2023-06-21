
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
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
            Text = string.Empty;
        }

        public ExcelConditionalFormattingNotContainsText(
          ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(eExcelConditionalFormattingRuleType.ContainsText, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.NotContains;
        }

        ExcelConditionalFormattingNotContainsText(ExcelConditionalFormattingNotContainsText copy) :base(copy)
        {
            Text = copy.Text;
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingNotContainsText(this);
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
                Formula = string.Format(
                  "ISERROR(SEARCH(\"{1}\",{0}))",
                  Address.Start.Address,
                  value.Replace("\"", "\"\""));
            }
        }

        public string FormulaReference
        {
            get
            {
                return Formula2;
            }
            set
            {
                _text = null;
                Formula2 = value;

                Formula = string.Format(
                    "ISERROR(SEARCH({1},{0}))",
                    Address.Start.Address,
                    value.Replace("\"", "\"\""));
            }
        }

        void UpdateFormula()
        {
            if (_text != null)
            {
                Formula = string.Format(
                    "ISERROR(SEARCH(\"{1}\",{0}))",
                    Address.Start.Address,
                    _text);
            }
            else if (Formula2 != null)
            {
                Formula = string.Format(
                    "ISERROR(SEARCH({1},{0}))",
                    Address.Start.Address,
                    Formula2);
            }
        }

        internal override bool IsExtLst
        {
            get
            {
                if (Formula2 != null)
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
