﻿
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

            Text = Formula.GetSubstringStoppingAtSymbol("NOT(ISERROR(SEARCH(\"".Length);
        }

        ExcelConditionalFormattingContainsText(ExcelConditionalFormattingContainsText copy) :base(copy)
        {
            ContainText = copy.ContainText;
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
                _FormulaReference = null;
                //TODO: Error check/Throw when formula does not follow this format and is a ContainsText.
                Formula = string.Format(
                  "NOT(ISERROR(SEARCH(\"{1}\",{0})))",
                  Address.Start.Address,
                  value.Replace("\"", "\"\""));
            }
        }

        string _FormulaReference = null;

        public string FormulaReference
        {
            get
            {
                return _FormulaReference;
            }
            set
            {
                Text = null;
                _FormulaReference = value;
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
            else if(_FormulaReference != null) 
            {
                Formula = string.Format(
                "NOT(ISERROR(SEARCH({1},{0})))",
                Address.Start.Address,
                _FormulaReference);
            }
        }

        public override ExcelAddress Address
        {
            get { return base.Address; }
            set { base.Address = value; UpdateFormula(); }
        }
    }
}
