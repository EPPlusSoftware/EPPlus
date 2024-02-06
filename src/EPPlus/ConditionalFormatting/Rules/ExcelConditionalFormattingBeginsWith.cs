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
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingBeginsWith : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingBeginsWith
    {
        /****************************************************************************************/

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingBeginsWith(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(
                eExcelConditionalFormattingRuleType.BeginsWith,
                address,
                priority,
                worksheet
                )
        {
            Operator = eExcelConditionalFormattingOperatorType.BeginsWith;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingBeginsWith(
          ExcelAddress address,
          ExcelWorksheet worksheet,
          XmlReader xr)
          : base(
                eExcelConditionalFormattingRuleType.BeginsWith,
                address,
                worksheet,
                xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.BeginsWith;
        }

        internal ExcelConditionalFormattingBeginsWith(ExcelConditionalFormattingBeginsWith copy, ExcelWorksheet newWs = null) : base(copy, newWs)
        {
            Operator = copy.Operator;
            Text = copy._text;
        }

        internal override ExcelConditionalFormattingRule Clone(ExcelWorksheet newWs = null)
        {
            return new ExcelConditionalFormattingBeginsWith(this, newWs);
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
                    "LEFT({0},LEN({1}))={1}",
                    Address.Start.Address,
                    value);
            }
        }

        public string Text {
            get { return _text; }
            set
            {
                _text = value;
                Formula2 = null;

                base.Formula = string.Format(
                  "LEFT({0},LEN(\"{1}\"))=\"{1}\"",
                  Address.Start.Address,
                  value.Replace("\"", "\"\""));
            }
        }

        void UpdateFormula()
        {
            if (_text != null)
            {
                base.Formula = string.Format(
                    "LEFT({0},LEN(\"{1}\"))=\"{1}\"",
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
            set { base.Address = value;
                if (value != null)
                {
                    UpdateFormula();
                }
            }
        }

        internal override bool ShouldApplyToCell(ExcelAddress address)
        {
            if (Address.Collide(address) != ExcelAddressBase.eAddressCollition.No)
            {
                if (_ws.Cells[Address.Start.Address].Value != null)
                {
                    var cellString = _ws.Cells[address.Start.Address].Value.ToString();

                    if (Formula2 != null)
                    {
                        return CultureInfo.CurrentCulture.CompareInfo.IsPrefix(cellString, Formula2);
                    }
                    else
                    {
                        return CultureInfo.CurrentCulture.CompareInfo.IsPrefix(cellString, _text);
                    }
                }
            }

            return false;
        }

        #endregion Constructors

        /****************************************************************************************/
    }
}
