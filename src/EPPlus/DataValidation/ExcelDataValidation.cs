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
 *************************************************************************************************/
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Utils;
using System;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Abstract base class for all Excel datavalidations. Contains functionlity which is common for all these different validation types.
    /// </summary>
    public abstract class ExcelDataValidation : IExcelDataValidation
    {
        /// <summary>
        /// Xml path for Formula1
        /// </summary>
        private readonly string _formula1Path = "d:formula1";
        private readonly string _formula1ExtLstPath = "x14:formula1/xm:f";

        /// <summary>
        /// Xml path for Formula2
        /// </summary>
        private readonly string _formula2Path = "d:formula2";
        private readonly string _formula2ExtLstPath = "x14:formula2/xm:f";

        public string Uid { get; internal set; }

        public ExcelAddress Address { get; internal set; }

        public virtual ExcelDataValidationType ValidationType { get; }

        string errorStyleString = null;
        /// <summary>
        /// Warning style
        /// </summary>
        public ExcelDataValidationWarningStyle ErrorStyle
        {
            get
            {
                if (!string.IsNullOrEmpty(errorStyleString))
                    return (ExcelDataValidationWarningStyle)Enum.Parse(typeof(ExcelDataValidationWarningStyle), errorStyleString);

                return ExcelDataValidationWarningStyle.undefined;
            }
            set
            {
                if (value == ExcelDataValidationWarningStyle.undefined)
                    errorStyleString = null;
                else
                    errorStyleString = value.ToString();
            }
        }

        public bool? AllowBlank { get; set; } = null;

        public bool? ShowInputMessage { get; set; } = null;

        public bool? ShowErrorMessage { get; set; } = null;

        public string ErrorTitle { get; set; } = null;

        public string Error { get; set; } = null;

        public string PromptTitle { get; set; } = null;

        public string Prompt { get; set; } = null;

        public virtual bool AllowsOperator { get { return true; } }


        internal string IFormula1 = null;
        internal string IFormula2 = null;
        string dvAddress = null;

        /// <summary>
        /// This method will validate the state of the validation
        /// </summary>
        /// <exception cref="InvalidOperationException">If the state breaks the rules of the validation</exception>
        public virtual void Validate()
        {
            var address = Address.Address;
            // validate Formula1
            if (string.IsNullOrEmpty(IFormula1) && !(AllowBlank ?? false))
            {
                throw new InvalidOperationException("Validation of " + address + " failed: Formula1 cannot be empty");
            }
        }

        public ExcelDataValidationAsType As { get; }

        /// <summary>
        /// Indicates whether this instance is stale, see https://github.com/EPPlusSoftware/EPPlus/wiki/Data-validation-Exceptions
        /// </summary>
        public bool IsStale { get; }

        string operatorString = null;
        /// <summary>
        /// Operator for comparison between the entered value and Formula/Formulas.
        /// </summary>
        public ExcelDataValidationOperator Operator
        {
            get
            {
                if (!string.IsNullOrEmpty(operatorString))
                {
                    return (ExcelDataValidationOperator)Enum.Parse(typeof(ExcelDataValidationOperator), operatorString);
                }
                return default(ExcelDataValidationOperator);
            }
            set
            {
                if ((ValidationType.Type == eDataValidationType.Any) || ValidationType.Type == eDataValidationType.List)
                {
                    throw new InvalidOperationException("The current validation type does not allow operator to be set");
                }
                //CheckIfStale();
                //SetXmlNodeString(_operatorPath, value.ToString());
            }
        }


        //virtual internal void Load(XmlReader xr)
        //{
        //    while (xr.Read())
        //    {
        //        if (xr.LocalName != "dataValidation") break;

        //        if (xr.NodeType == XmlNodeType.Element)
        //        {
        //            string validationType = xr.GetAttribute("type");
        //            string address = xr.GetAttribute("sqref");

        //        }
        //    }
        //}

        internal InternalValidationType InternalValidationType { get; set; } = InternalValidationType.DataValidation;

        internal void SetInternalValidationType(InternalValidationType type)
        {
            InternalValidationType = type;
        }

        protected ExcelDataValidation(string uid, string address)
        {
            Require.Argument(uid).IsNotNullOrEmpty("uid");
            Require.Argument(address).IsNotNullOrEmpty("address");

            Uid = uid;
            Address = new ExcelAddress(address);
        }

        protected ExcelDataValidation(XmlReader xr)
        {
            LoadXML(xr);
        }

        internal virtual void LoadXML(XmlReader xr)
        {
            string address = xr.GetAttribute("sqref");
            if (address == null)
                InternalValidationType = InternalValidationType.ExtLst;

            Uid = xr.GetAttribute("xr:uid");

            operatorString = xr.GetAttribute("operator");
            errorStyleString = xr.GetAttribute("errorStyle");

            AllowBlank = xr.GetAttribute("allowBlank") == "1" ? true : false;

            ShowInputMessage = xr.GetAttribute("showInputMessage") == "1" ? true : false;
            ShowErrorMessage = xr.GetAttribute("showErrorMessage") == "1" ? true : false;

            ErrorTitle = xr.GetAttribute("errorTitle");
            Error = xr.GetAttribute("error");

            PromptTitle = xr.GetAttribute("promptTitle");
            Prompt = xr.GetAttribute("prompt");

            ReadClassSpecificXmlNodes(xr);

            if (address == null && xr.ReadUntil(5, "sqref", "dataValidation", "extLst"))
            {
                address = xr.ReadString();
                if (address == null)
                {
                    throw new NullReferenceException($"Unable to locate ExtList adress for DataValidation with uid:{Uid}");
                }
            }

            Address = new ExcelAddress(address);
        }

        internal virtual void ReadClassSpecificXmlNodes(XmlReader xr)
        {

        }



        internal static string NewId()
        {
            return "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
        }

        internal void SetAddress(string address)
        {
            var dvAddress = AddressUtility.ParseEntireColumnSelections(address);
            dvAddress = address;
        }


    }
}

