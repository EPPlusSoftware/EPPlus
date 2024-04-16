
/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/
using System.Globalization;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingContainsText : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingContainsText
    {
        internal ExcelConditionalFormattingContainsText(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.ContainsText, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.ContainsText;
            Text = string.Empty;
        }

        internal ExcelConditionalFormattingContainsText(
          ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(eExcelConditionalFormattingRuleType.ContainsText, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.ContainsText;
        }

        ExcelConditionalFormattingContainsText(ExcelConditionalFormattingContainsText copy, ExcelWorksheet newWs = null) :base(copy, newWs)
        {
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingContainsText(this, newWs);
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
                if(Address != null)
                {
                    base.Formula = string.Format(
                      "NOT(ISERROR(SEARCH(\"{1}\",{0})))",
                      Address.Start.Address,
                      _text);
                }
                else
                {
                    base.Formula = string.Format(
                      "NOT(ISERROR(SEARCH(\"{1}\",{0})))",
                      "#REF!",
                      _text);
                }
            }
            else if(Formula2 != null) 
            {
                Formula = Formula2;
            }
        }
        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            if(Address.Collide(address) != ExcelAddressBase.eAddressCollition.No)
            {
                var val = _ws.Cells[address.Start.Address].Value;
                var stringValue = val == null ? "" : val.ToString();
                //Formula2 only filled if there's a cell or formula to apply a conditionalformat to.
                if (Formula2 != null)
                {
                    Formula = Formula2;
                    return CultureInfo.CurrentCulture.CompareInfo.IndexOf(stringValue, Formula2, CompareOptions.IgnoreCase) >= 0;
                }
                else if(_text != null)
                {
                    return CultureInfo.CurrentCulture.CompareInfo.IndexOf(stringValue, _text, CompareOptions.IgnoreCase) >= 0;
                }
            }

            return false;
        }

        public override ExcelAddress Address
        {
            get { return base.Address; }
            set { base.Address = value; UpdateFormula(); }
        }
    }
}
