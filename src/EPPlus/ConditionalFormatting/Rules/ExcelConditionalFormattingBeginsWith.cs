using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
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

            _containText = Text;
        }

        internal ExcelConditionalFormattingBeginsWith(ExcelConditionalFormattingBeginsWith copy) : base(copy)
        {
            Operator = copy.Operator;
            ContainText = copy.Text;
        }

        internal override ExcelConditionalFormattingRule Clone()
        {
            return new ExcelConditionalFormattingBeginsWith(this);
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
                Formula = string.Format(
                    "LEFT({0},LEN({1}))={1}",
                    Address.Start.Address,
                    value.Replace("\"", "\"\""));
                Formula2 = value;
            }
        }

        private string _containText = ""; 

        public string ContainText {
            get { return _containText; }
            set
            {
                _containText = value;
                Text = value;
                _formulaReference = null;
                Formula2 = null;

                Formula = string.Format(
                  "LEFT({0},LEN(\"{1}\"))=\"{1}\"",
                  Address.Start.Address,
                  value.Replace("\"", "\"\""));
            }
        }

        void UpdateFormula()
        {
            if (Text != null)
            {
                Formula = string.Format(
                    "LEFT({0},LEN(\"{1}\"))=\"{1}\"",
                    Address.Start.Address,
                    Text);
            }
            else if(_formulaReference != null)
            {
                Formula = string.Format(
                    "LEFT({0},LEN({1}))={1}",
                    Address.Start.Address,
                    _formulaReference);
            }
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
