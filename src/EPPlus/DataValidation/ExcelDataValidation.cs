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
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;

using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{

    /// <summary>
    /// Abstract base class for all Excel datavalidations. Contains functionlity which is common for all these different validation types.
    /// </summary>
    public abstract class ExcelDataValidation : IExcelDataValidation
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="uid">Id for validation</param>
        /// <param name="address">adress validation is applied to</param>
        /// <param name="ws">The worksheet</param>
        protected ExcelDataValidation(string uid, string address, ExcelWorksheet ws)
        {
            Require.Argument(uid).IsNotNullOrEmpty("uid");
            Require.Argument(address).IsNotNullOrEmpty("address");

            Uid = uid;
            _address = new ExcelDatavalidationAddress(CheckAndFixRangeAddress(address), this);
            _ws = ws;
        }

        /// <summary>
        /// Read-File Constructor
        /// </summary>
        /// <param name="xr"></param>
        /// <param name="ws">The worksheet</param>
        protected ExcelDataValidation(XmlReader xr, ExcelWorksheet ws)
        {
            LoadXML(xr);
            _ws = ws;
        }

        /// <summary>
        /// Copy-Constructor
        /// </summary>
        /// <param name="validation">Validation to copy from</param>
        /// <param name="ws">The worksheet</param>
        protected ExcelDataValidation(ExcelDataValidation validation,ExcelWorksheet ws)
        {
            Uid = validation.Uid;
            Address = validation.Address;
            ValidationType = validation.ValidationType;
            ErrorStyle = validation.ErrorStyle;
            AllowBlank = validation.AllowBlank;
            ShowInputMessage = validation.ShowInputMessage;
            ShowErrorMessage = validation.ShowErrorMessage;
            ErrorTitle = validation.ErrorTitle;
            Error = validation.Error;
            PromptTitle = validation.PromptTitle;
            Prompt = validation.Prompt;
            operatorString = validation.operatorString;
            _ws = ws;
        }

        internal ExcelWorksheet _ws;

        /// <summary>
        /// Uid of the data validation
        /// </summary>
        public string Uid { get; internal set; }

        ExcelDatavalidationAddress _address;

        /// <summary>
        /// Address of data validation
        /// </summary>
        public ExcelAddress Address { get { return _address; } internal set { _address = new ExcelDatavalidationAddress(value.Address, this); } }

        /// <summary>
        /// Validation type
        /// </summary>
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
                    return (ExcelDataValidationWarningStyle)Enum.Parse(typeof(ExcelDataValidationWarningStyle), errorStyleString, true);

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

        string imeModeString = null;
        /// <summary>
        ///  Mode for east-asian languages who use Input Method Editors(IME)
        /// </summary>
        public ExcelDataValidationImeMode ImeMode
        {
            get
            {
                if (string.IsNullOrEmpty(imeModeString))
                    return (ExcelDataValidationImeMode.NoControl);

                return (ExcelDataValidationImeMode) imeModeString.ToEnum<ExcelDataValidationImeMode>();
            }
            set
            {
                if (value == ExcelDataValidationImeMode.NoControl)
                    imeModeString = null;
                else
                    imeModeString = value.ToString();
            }
        }

        /// <summary>
        /// True if blanks should be allowed
        /// </summary>
        public bool? AllowBlank { get; set; } = null;

        /// <summary>
        /// True if input message should be shown
        /// </summary>
        public bool? ShowInputMessage { get; set; } = null;

        /// <summary>
        /// True if error message should be shown
        /// </summary>
        public bool? ShowErrorMessage { get; set; } = null;

        /// <summary>
        /// Title of error message box
        /// </summary>
        public string ErrorTitle { get; set; } = null;

        /// <summary>
        /// Error message box text
        /// </summary>
        public string Error { get; set; } = null;

        /// <summary>
        /// Title of the validation message box.
        /// </summary>
        public string PromptTitle { get; set; } = null;

        /// <summary>
        /// Text of the validation message box.
        /// </summary>
        public string Prompt { get; set; } = null;

        /// <summary>
        /// True if the current validation type allows operator.
        /// </summary>
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
        /// DEPRECATED as of Epplus 6.2.
        /// This as validations can no longer be stale since all attributes are now always fresh and held in the system.
        /// </summary>
        [Obsolete]
        public bool IsStale { get; } = false;

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
                    return (ExcelDataValidationOperator)Enum.Parse(typeof(ExcelDataValidationOperator), operatorString, true);
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

            var tempAddress = new ExcelAddress(address);
            string wsName = "";

            if (!string.IsNullOrEmpty(tempAddress.WorkSheetName))
            {
                address = address.Replace("'", "");

                wsName = tempAddress.WorkSheetName + "!";
                address = address.Replace(wsName, "");
            }

            address = ConvertUtil._invariantTextInfo.ToUpper(address);

            if (IsEntireColumn(address))
            {
                address = AddressUtility.ParseEntireColumnSelections(address);
            }
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

        /// <summary>
        /// Type to determine if extLst or not
        /// </summary>
        internal InternalValidationType InternalValidationType { get; set; } = InternalValidationType.DataValidation;


        /// <summary>
        /// Event method for changing internal type when referring to an external worksheet.
        /// </summary>
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

            imeModeString = xr.GetAttribute("imeMode");

            AllowBlank = xr.GetAttribute("allowBlank") == "1" || xr.GetAttribute("allowBlank") == "true" ? true : false;

            ShowInputMessage = xr.GetAttribute("showInputMessage") == "1" || xr.GetAttribute("showInputMessage") == "true" ? true : false;
            ShowErrorMessage = xr.GetAttribute("showErrorMessage") == "1" || xr.GetAttribute("showErrorMessage") == "true" ? true : false;

            ErrorTitle = xr.GetAttribute("errorTitle");
            Error = xr.GetAttribute("error");

            PromptTitle = xr.GetAttribute("promptTitle");
            Prompt = xr.GetAttribute("prompt");

            //Handle self-ending nodes
            var isSelfClosing = xr.IsEmptyElement;

            xr.Read();

            //Still needs to be done to "fill" formulas with null values even if empty
            ReadClassSpecificXmlNodes(xr);

            if (isSelfClosing == false)
            {
                //Can only happen if extLst
                if (address == null)
                {
                    address = xr.ReadString();
                    if (address == null)
                    {
                        throw new NullReferenceException($"Unable to locate ExtList adress for DataValidation with uid:{Uid}");
                    }
                    xr.Read();
                }
            }

            _address = new ExcelDatavalidationAddress
                (CheckAndFixRangeAddress(address)
                 .Replace(" ", ","), this);

            if(xr.LocalName == "dataValidations" || isSelfClosing)
            {
                return;
            }

            //Should only happen with 'legacy' files if for example an 'equals' validation has a formula2 node.
            if(xr.LocalName != "dataValidation")
            {
                var name = xr.Name;
                var type = xr.NodeType;
                //Since we handle self-closing nodes there must exist an endnode
                xr.ReadUntil("dataValidation");

                //This should not be logically possible unless file is corrupt.
                if(xr.LocalName != "dataValidation")
                {
                    throw new BadReadException($"Expected dataValidation end node could not be found." +
                        $"Last node name: '{name}' of type '{type}'");
                }
            }

            //Read to next dataValidation or dataValidations end node
            xr.Read();
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
            _address = new ExcelDatavalidationAddress(address, this);
        }

        /// <summary>
        /// Create a Deep-Copy of this validation.
        /// Note that one should also implement a separate clone() method casting to the child class
        /// </summary>
        internal abstract ExcelDataValidation GetClone();

        /// <summary>
        /// Create a Deep-Copy of this validation.
        /// Note that one should also implement a separate clone() method casting to the child class
        /// </summary>
        internal abstract ExcelDataValidation GetClone(ExcelWorksheet copy);
    }
}

