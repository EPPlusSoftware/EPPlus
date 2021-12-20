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
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Enum for available data validation types
    /// </summary>
    public enum eDataValidationType
    {
        /// <summary>
        /// Any value
        /// </summary>
        Any,
        /// <summary>
        /// Integer value
        /// </summary>
        Whole,
        /// <summary>
        /// Decimal values
        /// </summary>
        Decimal,
        /// <summary>
        /// List of values
        /// </summary>
        List,
        /// <summary>
        /// Text length validation
        /// </summary>
        TextLength,
        /// <summary>
        /// DateTime validation
        /// </summary>
        DateTime,
        /// <summary>
        /// Time validation
        /// </summary>
        Time,
        /// <summary>
        /// Custom validation
        /// </summary>
        Custom
    }

    internal static class DataValidationSchemaNames
    {
        public const string Any = "";
        public const string None = "none";
        public const string Whole = "whole";
        public const string Decimal = "decimal";
        public const string List = "list";
        public const string TextLength = "textLength";
        public const string Date = "date";
        public const string Time = "time";
        public const string Custom = "custom";
    }

    /// <summary>
    /// Types of datavalidation
    /// </summary>
    public class ExcelDataValidationType
    {
        private ExcelDataValidationType(eDataValidationType validationType, bool allowOperator, string schemaName)
        {
            Type = validationType;
            AllowOperator = allowOperator;
            SchemaName = schemaName;
        }

        /// <summary>
        /// Validation type
        /// </summary>
        public eDataValidationType Type
        {
            get;
            private set;
        }

        internal string SchemaName
        {
            get;
            private set;
        }

        /// <summary>
        /// This type allows operator to be set
        /// </summary>
        internal bool AllowOperator
        {

            get;
            private set;
        }

        /// <summary>
        /// Returns a validation type by <see cref="eDataValidationType"/>
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        internal static ExcelDataValidationType GetByValidationType(eDataValidationType type)
        {
            switch (type)
            {
                case eDataValidationType.Any:
                    return ExcelDataValidationType.Any;
                case eDataValidationType.Whole:
                    return ExcelDataValidationType.Whole;
                case eDataValidationType.List:
                    return ExcelDataValidationType.List;
                case eDataValidationType.Decimal:
                    return ExcelDataValidationType.Decimal;
                case eDataValidationType.TextLength:
                    return ExcelDataValidationType.TextLength;
                case eDataValidationType.DateTime:
                    return ExcelDataValidationType.DateTime;
                case eDataValidationType.Time:
                    return ExcelDataValidationType.Time;
                case eDataValidationType.Custom:
                    return ExcelDataValidationType.Custom;
                default:
                    throw new InvalidOperationException("Non supported Validationtype : " + type.ToString());
            }
        }

        internal static ExcelDataValidationType GetBySchemaName(string schemaName)
        {
            switch (schemaName)
            {
                case DataValidationSchemaNames.Any:
                case DataValidationSchemaNames.None:
                    return ExcelDataValidationType.Any;
                case DataValidationSchemaNames.Whole:
                    return ExcelDataValidationType.Whole;
                case DataValidationSchemaNames.Decimal:
                    return ExcelDataValidationType.Decimal;
                case DataValidationSchemaNames.List:
                    return ExcelDataValidationType.List;
                case DataValidationSchemaNames.TextLength:
                    return ExcelDataValidationType.TextLength;
                case DataValidationSchemaNames.Date:
                    return ExcelDataValidationType.DateTime;
                case DataValidationSchemaNames.Time:
                    return ExcelDataValidationType.Time;
                case DataValidationSchemaNames.Custom:
                    return ExcelDataValidationType.Custom;
                default:
                    throw new ArgumentException("Invalid schemaname: " + schemaName);
            }
        }

        /// <summary>
        /// Overridden Equals, compares on internal validation type
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ExcelDataValidationType))
            {
                return false;
            }
            return ((ExcelDataValidationType)obj).Type == Type;
        }

        /// <summary>
        /// Overrides GetHashCode()
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        
        private static ExcelDataValidationType _any;
        /// <summary>
        /// Any values
        /// </summary>
        public static ExcelDataValidationType Any
        {
            get
            {
                if (_any == null)
                {
                    _any = new ExcelDataValidationType(eDataValidationType.Any, false, DataValidationSchemaNames.Any);
                }
                return _any;
            }
        }

        /// <summary>
        /// Integer values
        /// </summary>
        private static ExcelDataValidationType _whole;
        /// <summary>
        /// Whole/Integer validation type.
        /// </summary>
        public static ExcelDataValidationType Whole
        {
            get
            {
                if (_whole == null)
                {
                    _whole = new ExcelDataValidationType(eDataValidationType.Whole, true, DataValidationSchemaNames.Whole);
                }
                return _whole;
            }
        }

        
        private static ExcelDataValidationType _list;
        /// <summary>
        /// List validation type, a provided list of allowed values
        /// </summary>
        public static ExcelDataValidationType List
        {
            get
            {
                if (_list == null)
                {
                    _list = new ExcelDataValidationType(eDataValidationType.List, false, DataValidationSchemaNames.List);
                }
                return _list;
            }
        }

        private static ExcelDataValidationType _decimal;
        /// <summary>
        /// Decimal validation type
        /// </summary>
        public static ExcelDataValidationType Decimal
        {
            get
            {
                if (_decimal == null)
                {
                    _decimal = new ExcelDataValidationType(eDataValidationType.Decimal, true, DataValidationSchemaNames.Decimal);
                }
                return _decimal;
            }
        }

        private static ExcelDataValidationType _textLength;
        /// <summary>
        /// Text length validation type
        /// </summary>
        public static ExcelDataValidationType TextLength
        {
            get
            {
                if (_textLength == null)
                {
                    _textLength = new ExcelDataValidationType(eDataValidationType.TextLength, true, DataValidationSchemaNames.TextLength);
                }
                return _textLength;
            }
        }

        private static ExcelDataValidationType _dateTime;
        /// <summary>
        ///  Time validation type
        /// </summary>
        public static ExcelDataValidationType DateTime
        {
            get
            {
                if (_dateTime == null)
                {
                    _dateTime = new ExcelDataValidationType(eDataValidationType.DateTime, true, DataValidationSchemaNames.Date);
                }
                return _dateTime;
            }
        }

        private static ExcelDataValidationType _time;
        /// <summary>
        /// Time validation type
        /// </summary>
        public static ExcelDataValidationType Time
        {
            get
            {
                if (_time == null)
                {
                    _time = new ExcelDataValidationType(eDataValidationType.Time, true, DataValidationSchemaNames.Time);
                }
                return _time;
            }
        }

        private static ExcelDataValidationType _custom;
        /// <summary>
        /// Custom validation type
        /// </summary>
        public static ExcelDataValidationType Custom
        {
            get
            {
                if (_custom == null)
                {
                    _custom = new ExcelDataValidationType(eDataValidationType.Custom, true, DataValidationSchemaNames.Custom);
                }
                return _custom;
            }
        }
    }
}
