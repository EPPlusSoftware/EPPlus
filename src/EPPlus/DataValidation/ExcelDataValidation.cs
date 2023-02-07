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
using OfficeOpenXml.DataValidation.Events;
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

        /// <summary>
        /// This method will validate the state of the validation
        /// </summary>
        /// <exception cref="InvalidOperationException">If the state breaks the rules of the validation</exception>
        public virtual void Validate()
        {

        }

        ExcelDataValidationAsType _as = null;
        /// <summary>
        /// Us this property to case <see cref="IExcelDataValidation"/>s to its subtypes
        /// </summary>
        public ExcelDataValidationAsType As
        {
            get
            {
                if (_as == null)
                {
                    _as = new ExcelDataValidationAsType(this);
                }
                return _as;
            }
        }

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
                operatorString = value.ToString();
            }
        }

        private string CheckAndFixRangeAddress(string address)
        {
            if (address.Contains(","))
            {
                throw new FormatException("Multiple addresses may not be commaseparated, use space instead");
            }
            address = ConvertUtil._invariantTextInfo.ToUpper(address);

            if (IsEntireColumn(address))
            {
                address = AddressUtility.ParseEntireColumnSelections(address);
            }
            //if (Regex.IsMatch(address, @"[A-Z]+:[A-Z]+"))
            //{
            //    address = AddressUtility.ParseEntireColumnSelections(address);
            //}
            return address;
        }

        bool IsEntireColumn(string address)
        {
            bool hasColon = false;
            foreach (char c in address)
            {
                if (((c >= 'A') && (c <= 'Z')) || c == ':')
                {
                    if (c == ':')
                    {
                        hasColon = true;
                    }
                    continue;
                }
                else
                {
                    return false;
                }
            }

            if (hasColon)
            {
                return true;
            }
            else
            {
                return false;
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
            Address = new ExcelAddress(CheckAndFixRangeAddress(address));
        }



        protected ExcelDataValidation(XmlReader xr)
        {
            LoadXML(xr);
        }

        protected Action<OnFormulaChangedEventArgs> OnFormulaChanged => (e) =>
        {
            if (e.isExt)
            {
                InternalValidationType = InternalValidationType.ExtLst;
            }
        };

        internal virtual void LoadXML(XmlReader xr)
        {
            string address = xr.GetAttribute("sqref");
            if (address == null)
                InternalValidationType = InternalValidationType.ExtLst;

            Uid = string.IsNullOrEmpty(xr.GetAttribute("xr:uid")) ? NewId() : xr.GetAttribute("xr:uid");

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

            Address = new ExcelAddress(CheckAndFixRangeAddress(address));
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
            Address = new ExcelAddress(address);
        }
    }
}

