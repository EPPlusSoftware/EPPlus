
using OfficeOpenXml.DataValidation.Contracts;

namespace OfficeOpenXml.DataValidation
{
    public abstract class NewDataValidation : IExcelDataValidation
    {
        private ExcelWorksheet worksheet;
        private string address;
        private object value;

        public string Uid { get; }

        public ExcelAddress Address { get; }

        public ExcelDataValidationType ValidationType { get; }

        public ExcelDataValidationWarningStyle ErrorStyle { get; set; }

        public bool? AllowBlank { get; set; }

        public bool? ShowInputMessage { get; set; }

        public bool? ShowErrorMessage { get; set; }

        public string ErrorTitle { get; set; }

        public string Error { get; set; }

        public string PromptTitle { get; set; }

        public string Prompt { get; set; }

        public bool AllowsOperator { get; }

        public void Validate()
        {

        }

          

        //private void ValidateAddress(string address, IExcelDataValidation validatingValidation)
        //{
        //    Require.Argument(address).IsNotNullOrEmpty("address");

        //    if (!InternalValidationEnabled) return;

        //    // ensure that the new address does not collide with an existing validation.
        //    var newAddress = new ExcelAddress(address);
        //    if (_validations.Count > 0)
        //    {
        //        foreach (var validation in _validations)
        //        {
        //            if (validatingValidation != null && validatingValidation == validation)
        //            {
        //                continue;
        //            }
        //            var result = validation.Address.Collide(newAddress);
        //            if (result != ExcelAddressBase.eAddressCollition.No)
        //            {
        //                throw new InvalidOperationException(string.Format("The address ({0}) collides with an existing validation ({1})", address, validation.Address.Address));
        //            }
        //        }
        //    }
        //}

        public ExcelDataValidationAsType As { get; }

        /// <summary>
        /// Indicates whether this instance is stale, see https://github.com/EPPlusSoftware/EPPlus/wiki/Data-validation-Exceptions
        /// </summary>
        public bool IsStale { get; }

        protected NewDataValidation(ExcelWorksheet worksheet, string uid, string address, object value)
        {
            this.worksheet = worksheet;
            Uid = uid;
            this.address = address;
            this.value = value;
        }

            //private const string ItemElementNodeName = "d:dataValidation";
            //private const string ExtLstElementNodeName = "x14:dataValidation";
            //private readonly string _uidPath = "@xr:uid";
            //private readonly string _errorStylePath = "@errorStyle";
            //private readonly string _errorTitlePath = "@errorTitle";
            //private readonly string _errorPath = "@error";
            //private readonly string _promptTitlePath = "@promptTitle";
            //private readonly string _promptPath = "@prompt";
            //private readonly string _operatorPath = "@operator";
            //private readonly string _showErrorMessagePath = "@showErrorMessage";
            //private readonly string _showInputMessagePath = "@showInputMessage";
            //private readonly string _typeMessagePath = "@type";
            //private readonly string _sqrefPath = "@sqref";
            //private readonly string _sqrefPathExt = "xm:sqref";
            //private readonly string _allowBlankPath = "@allowBlank";
            ///// <summary>
            ///// Xml path for Formula1
            ///// </summary>
            //private readonly string _formula1Path = "d:formula1";
            //private readonly string _formula1ExtLstPath = "x14:formula1/xm:f";
            ///// <summary>
            ///// Xml path for Formula2
            ///// </summary>
            //private readonly string _formula2Path = "d:formula2";
            //private readonly string _formula2ExtLstPath = "x14:formula2/xm:f";

            //public NewDataValidation() 
            //{

            //}
        }
}
