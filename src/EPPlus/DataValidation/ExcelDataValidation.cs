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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.Utils;
using System.Xml;
using System.Text.RegularExpressions;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.DataValidation.Contracts;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Abstract base class for all Excel datavalidations. Contains functionlity which is common for all these different validation types.
    /// </summary>
    public abstract class ExcelDataValidation : XmlHelper, IExcelDataValidation
    {
        private const string ItemElementNodeName = "d:dataValidation";
        private const string ExtLstElementNodeName = "x14:dataValidation";
        private readonly string _uidPath = "@xr:uid";
        private readonly string _errorStylePath = "@errorStyle";
        private readonly string _errorTitlePath = "@errorTitle";
        private readonly string _errorPath = "@error";
        private readonly string _promptTitlePath = "@promptTitle";
        private readonly string _promptPath = "@prompt";
        private readonly string _operatorPath = "@operator";
        private readonly string _showErrorMessagePath = "@showErrorMessage";
        private readonly string _showInputMessagePath = "@showInputMessage";
        private readonly string _typeMessagePath = "@type";
        private readonly string _sqrefPath = "@sqref";
        private readonly string _sqrefPathExt = "xm:sqref";
        private readonly string _allowBlankPath = "@allowBlank";
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

        internal ExcelDataValidation(ExcelWorksheet worksheet, string uid, string address, ExcelDataValidationType validationType, InternalValidationType internalValidationType = InternalValidationType.DataValidation)
            : this(worksheet, uid, address, validationType, null, internalValidationType)
        { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet">worksheet that owns the validation</param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="itemElementNode">Xml top node (dataValidations)</param>
        /// <param name="validationType">Data validation type</param>
        /// <param name="address">address for data validation</param>
        /// <param name="internalValidationType">If the datavalidation is internal or in the extLst element</param>
        internal ExcelDataValidation(ExcelWorksheet worksheet, string uid, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, InternalValidationType internalValidationType = InternalValidationType.DataValidation)
            : this(worksheet, uid, address, validationType, itemElementNode, null, internalValidationType)
        {

        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet">worksheet that owns the validation</param>
        /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
        /// <param name="itemElementNode">Xml top node (dataValidations) when importing xml</param>
        /// <param name="validationType">Data validation type</param>
        /// <param name="address">address for data validation</param>
        /// <param name="namespaceManager">Xml Namespace manager</param>
        /// <param name="internalValidationType"><see cref="InternalValidationType"/></param>
        internal ExcelDataValidation(ExcelWorksheet worksheet, string uid, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager, InternalValidationType internalValidationType = InternalValidationType.DataValidation)
            : base(namespaceManager != null ? namespaceManager : worksheet.NameSpaceManager)
        {
            Require.Argument(uid).IsNotNullOrEmpty("uid");
            Require.Argument(address).IsNotNullOrEmpty("address");
            InternalValidationType = internalValidationType;
            InitNodeOrder(validationType);
            address = CheckAndFixRangeAddress(address);
            if (itemElementNode == null)
            {
                TopNode = worksheet.WorksheetXml.SelectSingleNode(GetTopNodeName(), worksheet.NameSpaceManager);
                itemElementNode = CreateNode(GetItemElementNodeName(), false, true);
                TopNode.AppendChild(itemElementNode);
            }
            TopNode = itemElementNode;
            ValidationType = validationType;
            Uid = uid;
            Address = new ExcelAddress(address);
            
        }

        internal InternalValidationType InternalValidationType { get; private set; } = InternalValidationType.DataValidation;

        internal virtual void RegisterFormulaListener(DataValidationFormulaListener listener)
        {

        }

        private string GetSqRefPath()
        {
            return InternalValidationType == InternalValidationType.DataValidation ? _sqrefPath : _sqrefPathExt;
        }

        private string GetItemElementNodeName()
        {
            return InternalValidationType == InternalValidationType.DataValidation ? ItemElementNodeName : ExtLstElementNodeName;
        }

        private string GetTopNodeName()
        {
            return InternalValidationType == InternalValidationType.DataValidation ? "//d:dataValidations" : "//d:extLst/d:ext/x14:dataValidations";
        }

        private void InitNodeOrder(ExcelDataValidationType validationType)
        {
            // set schema node order
            if(validationType == ExcelDataValidationType.List || validationType == ExcelDataValidationType.Custom)
            {
                if(InternalValidationType == InternalValidationType.DataValidation)
                {
                    SchemaNodeOrder = new string[]{
                        "uid",
                        "type",
                        "errorStyle",
                        "allowBlank",
                        "showInputMessage",
                        "showErrorMessage",
                        "errorTitle",
                        "error",
                        "promptTitle",
                        "prompt",
                        "sqref",
                        "formula1"
                    };
                }
                else
                {
                    SchemaNodeOrder = new string[]{
                        "uid",
                        "type",
                        "errorStyle",
                        "allowBlank",
                        "showInputMessage",
                        "showErrorMessage",
                        "errorTitle",
                        "error",
                        "promptTitle",
                        "prompt",
                        "formula1",
                        "sqref"
                    };
                }
            }
            else
            {
                SchemaNodeOrder = new string[]{
                    "uid",
                    "type",
                    "errorStyle",
                    "operator",
                    "allowBlank",
                    "showInputMessage",
                    "showErrorMessage",
                    "errorTitle",
                    "error",
                    "promptTitle",
                    "prompt",
                    "sqref",
                    "formula1",
                    "formula2"
                };
            }
            
        }

        private string CheckAndFixRangeAddress(string address)
        {
            if (address.Contains(','))
            {
                throw new FormatException("Multiple addresses may not be commaseparated, use space instead");
            }
            address = ConvertUtil._invariantTextInfo.ToUpper(address);
            if (Regex.IsMatch(address, @"[A-Z]+:[A-Z]+"))
            {
                address = AddressUtility.ParseEntireColumnSelections(address);
            }
            return address;
        }

        private void SetNullableBoolValue(string path, bool? val)
        {
            if (val.HasValue)
            {
                SetXmlNodeBool(path, val.Value);
            }
            else
            {
                DeleteNode(path);
            }
        }

        /// <summary>
        /// This method will validate the state of the validation
        /// </summary>
        /// <exception cref="InvalidOperationException">If the state breaks the rules of the validation</exception>
        public virtual void Validate()
        {
            var address = Address.Address;
            // validate Formula1
            if (string.IsNullOrEmpty(Formula1Internal))
            {
                throw new InvalidOperationException("Validation of " + address + " failed: Formula1 cannot be empty");
            }
        }

        internal void Delete()
        {
            DeleteTopNode();
        }

        #region Public properties

        /// <summary>
        /// True if the validation type allows operator to be set.
        /// </summary>
        public bool AllowsOperator
        {
            get
            {
                return ValidationType.AllowOperator;
            }
        }

        public string Uid
        {
            get
            {
                return GetXmlNodeString(_uidPath);
            }
            set
            {
                if (string.IsNullOrEmpty(value)) throw new ArgumentNullException("Uid");
                var uid = value.TrimStart('{').TrimEnd('}');
                SetXmlNodeString(_uidPath, "{" + uid + "}");
            }
        }

        /// <summary>
        /// Address of data validation
        /// </summary>
        public ExcelAddress Address
        {
            get
            {
                return new ExcelAddress(GetXmlNodeString(GetSqRefPath()).Replace(" ", ","));
            }
            private set
            {
                SetAddress(value.Address);
            }
        }
        /// <summary>
        /// Validation type
        /// </summary>
        public ExcelDataValidationType ValidationType
        {
            get
            {
                var typeString = GetXmlNodeString(_typeMessagePath);
                return ExcelDataValidationType.GetBySchemaName(typeString);
            }
            private set
            {
                SetXmlNodeString(_typeMessagePath, value.SchemaName, true);
            }
        }

        /// <summary>
        /// Operator for comparison between the entered value and Formula/Formulas.
        /// </summary>
        public ExcelDataValidationOperator Operator
        {
            get
            {
                var operatorString = GetXmlNodeString(_operatorPath);
                if (!string.IsNullOrEmpty(operatorString))
                {
                    return (ExcelDataValidationOperator)Enum.Parse(typeof(ExcelDataValidationOperator), operatorString);
                }
                return default(ExcelDataValidationOperator);
            }
            set
            {
                if (!ValidationType.AllowOperator)
                {
                    throw new InvalidOperationException("The current validation type does not allow operator to be set");
                }
                SetXmlNodeString(_operatorPath, value.ToString());
            }
        }

        /// <summary>
        /// Warning style
        /// </summary>
        public ExcelDataValidationWarningStyle ErrorStyle
        {
            get
            {
                var errorStyleString = GetXmlNodeString(_errorStylePath);
                if (!string.IsNullOrEmpty(errorStyleString))
                {
                    return (ExcelDataValidationWarningStyle)Enum.Parse(typeof(ExcelDataValidationWarningStyle), errorStyleString);
                }
                return ExcelDataValidationWarningStyle.undefined;
            }
            set
            {
                if (value == ExcelDataValidationWarningStyle.undefined)
                {
                    DeleteNode(_errorStylePath);
                }
                else
                {
                    SetXmlNodeString(_errorStylePath, value.ToString());
                }
            }
        }

        /// <summary>
        /// True if blanks should be allowed
        /// </summary>
        public bool? AllowBlank
        {
            get
            {
                return GetXmlNodeBoolNullable(_allowBlankPath);
            }
            set
            {
                SetNullableBoolValue(_allowBlankPath, value);
            }
        }

        /// <summary>
        /// True if input message should be shown
        /// </summary>
        public bool? ShowInputMessage
        {
            get
            {
                return GetXmlNodeBoolNullable(_showInputMessagePath);
            }
            set
            {
                SetNullableBoolValue(_showInputMessagePath, value);
            }
        }

        /// <summary>
        /// True if error message should be shown
        /// </summary>
        public bool? ShowErrorMessage
        {
            get
            {
                return GetXmlNodeBoolNullable(_showErrorMessagePath);
            }
            set
            {
                SetNullableBoolValue(_showErrorMessagePath, value);
            }
        }

        /// <summary>
        /// Title of error message box
        /// </summary>
        public string ErrorTitle
        {
            get
            {
                return GetXmlNodeString(_errorTitlePath);
            }

            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    DeleteNode(_errorTitlePath);
                }
                else
                {
                    SetXmlNodeString(_errorTitlePath, value.ToString());
                }
            }
        }

        /// <summary>
        /// Error message box text
        /// </summary>
        public string Error
        {
            get
            {
                return GetXmlNodeString(_errorPath);
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    DeleteNode(_errorPath);
                }
                else
                {
                    SetXmlNodeString(_errorPath, value.ToString());
                }
            }
        }

        /// <summary>
        /// Title of the validation message box.
        /// </summary>
        public string PromptTitle
        {
            get
            {
                return GetXmlNodeString(_promptTitlePath);
            }
            set
            {
                SetXmlNodeString(_promptTitlePath, value);
            }
        }

        /// <summary>
        /// Text of the validation message box.
        /// </summary>
        public string Prompt
        {
            get
            {
                return GetXmlNodeString(_promptPath);
            }
            set
            {
                SetXmlNodeString(_promptPath, value);
            }
        }

        /// <summary>
        /// Formula 1
        /// </summary>
        protected string Formula1Internal
        {
            get
            {
                return GetXmlNodeString(GetFormula1Path());
            }
        }

        /// <summary>
        /// Formula 2
        /// </summary>
        protected string Formula2Internal
        {
            get
            {
                return GetXmlNodeString(GetFormula2Path());
            }
        }

        #endregion

        #region Internal properties

        internal static string NewId()
        {
            return "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
        }
        internal bool IsExtLst { get; set; }

        ExcelDataValidationAsType _as = null;
        /// <summary>
        /// Us this property to case <see cref="IExcelDataValidation"/>s to its subtypes
        /// </summary>
        public ExcelDataValidationAsType As
        {
            get
            {
                if(_as == null)
                {
                    _as = new ExcelDataValidationAsType(this);
                }
                return _as;
            }
        }
        #endregion

        /// <summary>
        /// Sets the value to the supplied path
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="val">The value to set</param>
        /// <param name="path">xml path</param>
        protected void SetValue<T>(Nullable<T> val, string path)
            where T : struct
        {
            if (!val.HasValue)
            {
                DeleteNode(path);
            }
            var stringValue = val.Value.ToString().Replace(',', '.');
            SetXmlNodeString(path, stringValue);
        }

        protected string GetFormula1Path()
        {
            return InternalValidationType == InternalValidationType.DataValidation ? _formula1Path : _formula1ExtLstPath;
        }

        protected string GetFormula2Path()
        {
            return InternalValidationType == InternalValidationType.DataValidation ? _formula2Path : _formula2ExtLstPath;
        }

        internal void SetAddress(string address)
        {
            var dvAddress = AddressUtility.ParseEntireColumnSelections(address);
            SetXmlNodeString(GetSqRefPath(), dvAddress.Replace(",", " "));
            
        }
    }
}
