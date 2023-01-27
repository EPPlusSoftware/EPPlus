using OfficeOpenXml.DataValidation.Contracts;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Provides a simple way to type cast a data validation object to its actual class.
    /// </summary>
    public class ExcelDataValidationAsType
    {
        IExcelDataValidation _validation;
        internal ExcelDataValidationAsType(IExcelDataValidation validation)
        {
            _validation = validation;
        }

        /// <summary>
        /// Converts the data validation object to it's implementing class or any of the abstract classes/interfaces inheriting the <see cref="IExcelDataValidation"/> interface.        
        /// </summary>
        /// <typeparam name="T">The type of datavalidation object. T must be inherited from <see cref="IExcelDataValidation"/></typeparam>
        /// <returns>An instance of <typeparamref name="T"/> or null if type casting fails.</returns>
        public T Type<T>() where T : IExcelDataValidation
        {
            if (_validation is T t)
            {
                return t;
            }
            return default;
        }

        /// <summary>
        /// Returns the data validation object as <see cref="IExcelDataValidationList"/>
        /// </summary>
        /// <returns>The data validation as an <see cref="IExcelDataValidationList"/> or null if typecasting fails</returns>
        public IExcelDataValidationList ListValidation
        {
            get
            {
                return _validation as IExcelDataValidationList;
            }
        }

        /// <summary>
        /// Returns the data validation object as <see cref="IExcelDataValidationInt"/>
        /// </summary>
        /// <returns>The data validation as an <see cref="IExcelDataValidationInt"/> or null if typecasting fails</returns>
        public IExcelDataValidationInt IntegerValidation
        {
            get
            {
                return _validation as IExcelDataValidationInt;
            }
        }

        /// <summary>
        /// Returns the data validation object as <see cref="IExcelDataValidationDateTime"/>
        /// </summary>
        /// <returns>The data validation as an <see cref="IExcelDataValidationDateTime"/> or null if typecasting fails</returns>
        public IExcelDataValidationDateTime DateTimeValidation
        {
            get
            {
                return _validation as IExcelDataValidationDateTime;
            }
        }

        /// <summary>
        /// Returns the data validation object as <see cref="IExcelDataValidationTime"/>
        /// </summary>
        /// <returns>The data validation as an <see cref="IExcelDataValidationTime"/> or null if typecasting fails</returns>
        public IExcelDataValidationTime TimeValidation
        {
            get
            {
                return _validation as IExcelDataValidationTime;
            }
        }

        /// <summary>
        /// Returns the data validation object as <see cref="IExcelDataValidationDecimal"/>
        /// </summary>
        /// <returns>The data validation as an <see cref="IExcelDataValidationDecimal"/> or null if typecasting fails</returns>
        public IExcelDataValidationDecimal DecimalValidation
        {
            get
            {
                return _validation as IExcelDataValidationDecimal;
            }
        }

        /// <summary>
        /// Returns the data validation object as <see cref="IExcelDataValidationAny"/>
        /// </summary>
        /// <returns>The data validation as an <see cref="IExcelDataValidationAny"/> or null if typecasting fails</returns>
        public IExcelDataValidationAny AnyValidation
        {
            get
            {
                return _validation as IExcelDataValidationAny;
            }
        }

        /// <summary>
        /// Returns the data validation object as <see cref="IExcelDataValidationCustom"/>
        /// </summary>
        /// <returns>The data validation as an <see cref="IExcelDataValidationCustom"/> or null if typecasting fails</returns>
        public IExcelDataValidationCustom CustomValidation
        {
            get
            {
                return _validation as IExcelDataValidationCustom;
            }
        }
    }
}
