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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// <para>
    /// Collection of <see cref="ExcelDataValidation"/>. This class is providing the API for EPPlus data validation.
    /// </para>
    /// <para>
    /// The public methods of this class (Add[...]Validation) will create a datavalidation entry in the worksheet. When this
    /// validation has been created changes to the properties will affect the workbook immediately.
    /// </para>
    /// <para>
    /// Each type of validation has either a formula or a typed value/values, except for custom validation which has a formula only.
    /// </para>
    /// <code>
    /// // Add a date time validation
    /// var validation = worksheet.DataValidation.AddDateTimeValidation("A1");
    /// // set validation properties
    /// validation.ShowErrorMessage = true;
    /// validation.ErrorTitle = "An invalid date was entered";
    /// validation.Error = "The date must be between 2011-01-31 and 2011-12-31";
    /// validation.Prompt = "Enter date here";
    /// validation.Formula.Value = DateTime.Parse("2011-01-01");
    /// validation.Formula2.Value = DateTime.Parse("2011-12-31");
    /// validation.Operator = ExcelDataValidationOperator.between;
    /// </code>
    /// </summary>
    public class ExcelDataValidationCollection : IEnumerable<ExcelDataValidation>
    {
        private List<ExcelDataValidation> _validations = new List<ExcelDataValidation>();
        private ExcelWorksheet _worksheet = null;
        private RangeDictionary<ExcelDataValidation> _validationsRD = new RangeDictionary<ExcelDataValidation>();

        internal ExcelDataValidationCollection(ExcelWorksheet worksheet)
        {
            InternalValidationEnabled = true;
            _worksheet = worksheet;
        }

        internal ExcelDataValidationCollection(XmlReader xr, ExcelWorksheet worksheet)
            : this(worksheet)
        {
            InternalValidationEnabled = true;
            ReadDataValidations(xr);
        }
        /// <summary>
        /// Read data validation from xml via xr reader
        /// </summary>
        public void ReadDataValidations(XmlReader xr)
        {
            while (xr.Read())
            {
                if (xr.LocalName != "dataValidation")
                {
                    xr.Read(); //Read beyond the end element
                    break;
                }

                if (xr.NodeType == XmlNodeType.Element)
                {
                    var validation = ExcelDataValidationFactory.Create(xr);
                    _validations.Add(validation);
                    _validationsRD.Add(validation.Address._fromRow, validation.Address._fromCol,
                                       validation.Address._toRow, validation.Address._toCol, validation);
                }
            }
        }

        internal bool HasValidationType(InternalValidationType type)
        {
            if (Count != 0)
            {
                for (int i = 0; i < Count; i++)
                {
                    if (_validations[i].InternalValidationType == type)
                    {
                        return true;
                    }
                }
                return false;
            }
            else
            {
                return false;
            }
        }

        int GetCount(InternalValidationType type)
        {
            int validationCount = 0;
            for (int i = 0; i < Count; i++)
            {
                if (_validations[i].InternalValidationType == type)
                {
                    validationCount++;
                }
            }
            return validationCount;
        }

        internal int GetNonExtLstCount()
        {
            return GetCount(InternalValidationType.DataValidation);
        }


        internal int GetExtLstCount()
        {
            return GetCount(InternalValidationType.ExtLst);
        }

        private void OnValidationCountChanged()
        {

        }

        /// <summary>
        /// Validates address - not empty, collisions
        /// </summary>
        /// <param name="address"></param>
        /// <param name="validatingValidation"></param>
        private void ValidateAddress(string address, IExcelDataValidation validatingValidation)
        {
            Require.Argument(address).IsNotNullOrEmpty("address");

            if (!InternalValidationEnabled) return;

            // ensure that the new address does not collide with an existing validation.
            var newAddress = new ExcelAddress(address);
            if (_validations.Count > 0)
            {
                foreach (var validation in _validations)
                {
                    if (validatingValidation != null && validatingValidation == validation)
                    {
                        continue;
                    }
                    var result = validation.Address.Collide(newAddress);
                    if (result != ExcelAddressBase.eAddressCollition.No)
                    {
                        throw new InvalidOperationException(string.Format("The address ({0}) collides with an existing validation ({1})", address, validation.Address.Address));
                    }
                }
            }
        }

        /// <summary>
        /// Validates all data validations.
        /// </summary>
        internal void ValidateAll()
        {
            if (!InternalValidationEnabled) return;

            foreach (var validation in _validations)
            {
                validation.Validate();
            }
        }

        internal void AddCopyOfDataValidation(string address, ExcelDataValidation dv)
        {
            _validations.Add(ExcelDataValidationFactory.CloneWithNewAdress(address, dv));
        }

        /// <summary>
        /// Adds a <see cref="ExcelDataValidationAny"/> to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public IExcelDataValidationAny AddAnyValidation(string address)
        {
            var validation = new ExcelDataValidationAny(ExcelDataValidation.NewId(), address);
            return (IExcelDataValidationAny)AddValidation(address, validation);
        }

        /// <summary>
        /// Adds an <see cref="IExcelDataValidationInt"/> to the worksheet. Whole means that the only accepted values
        /// are integer values.
        /// </summary>
        /// <param name="address">the range/address to validate</param>
        public IExcelDataValidationInt AddIntegerValidation(string address)
        {
            var validation = new ExcelDataValidationInt(ExcelDataValidation.NewId(), address, _worksheet.Name);
            return (IExcelDataValidationInt)AddValidation(address, validation);
        }

        /// <summary>
        /// Adds an <see cref="IExcelDataValidationInt"/> regarding text length to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        public IExcelDataValidationInt AddTextLengthValidation(string address)
        {
            var validation = new ExcelDataValidationInt(ExcelDataValidation.NewId(), address, _worksheet.Name, true);
            return (IExcelDataValidationInt)AddValidation(address, validation);
        }

        /// <summary>
        /// Addes an <see cref="IExcelDataValidationDecimal"/> to the worksheet. The only accepted values are
        /// decimal values.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        public IExcelDataValidationDecimal AddDecimalValidation(string address)
        {
            var validation = new ExcelDataValidationDecimal(ExcelDataValidation.NewId(), address, _worksheet.Name);
            return (IExcelDataValidationDecimal)AddValidation(address, validation);
        }

        /// <summary>
        /// Adds an <see cref="IExcelDataValidationList"/> to the worksheet. The accepted values are defined
        /// in a list.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        public IExcelDataValidationList AddListValidation(string address)
        {
            var validation = new ExcelDataValidationList(ExcelDataValidation.NewId(), address, _worksheet.Name);
            return (IExcelDataValidationList)AddValidation(address, validation);
        }

        /// <summary>
        /// Adds an <see cref="IExcelDataValidationDateTime"/> to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        public IExcelDataValidationDateTime AddDateTimeValidation(string address)
        {
            var validation = new ExcelDataValidationDateTime(ExcelDataValidation.NewId(), address, _worksheet.Name);
            return (IExcelDataValidationDateTime)AddValidation(address, validation);
        }

        /// <summary>
        /// Adds an <see cref="IExcelDataValidationDateTime"/> to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        public IExcelDataValidationTime AddTimeValidation(string address)
        {
            var validation = new ExcelDataValidationTime(ExcelDataValidation.NewId(), address, _worksheet.Name);
            return (IExcelDataValidationTime)AddValidation(address, validation);
        }

        /// <summary>
        /// Adds a <see cref="ExcelDataValidationCustom"/> to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        public IExcelDataValidationCustom AddCustomValidation(string address)
        {
            var validation = new ExcelDataValidationCustom(ExcelDataValidation.NewId(), address, _worksheet.Name);
            return (IExcelDataValidationCustom)AddValidation(address, validation);
        }

        private ExcelDataValidation AddValidation(string address, ExcelDataValidation validation)
        {
            var internalAddress = new ExcelAddress(address);

            if (_validationsRD.Exists(internalAddress._fromRow, internalAddress._fromCol, internalAddress._toRow, internalAddress._toCol))
            {
                throw new InvalidOperationException($"A DataValidation already exists at {address}");
            }

            _validations.Add(validation);
            _validationsRD.Add(internalAddress._fromRow, internalAddress._fromCol, internalAddress._toRow, internalAddress._toCol, validation);

            return validation;
        }

        /// <summary>
        /// Number of validations
        /// </summary>3
        public int Count
        {
            get { return _validations.Count; }
        }

        /// <summary>
        /// Epplus validates that all data validations are consistend and valid
        /// when they are added and when a workbook is saved. Since this takes some
        /// resources, it can be disabled for improve performance. 
        /// </summary>
        public bool InternalValidationEnabled
        {
            get;
            set;
        }

        /// <summary>
        /// Index operator, returns by 0-based index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelDataValidation this[int index]
        {
            get { return _validations[index]; }
            set { _validations[index] = value; }
        }

        /// <summary>
        /// Index operator, returns a data validation which address partly or exactly matches the searched address.
        /// </summary>
        /// <param name="address">A cell address or range</param>
        /// <returns>A <see cref="ExcelDataValidation"/> or null if no match</returns>
        public IExcelDataValidation this[string address]
        {
            get
            {
                var searchedAddress = new ExcelAddress(address);
                return _validations.Find(x => x.Address.Collide(searchedAddress) != ExcelAddressBase.eAddressCollition.No);
            }
        }

        /// <summary>
        /// Returns all validations that matches the supplied predicate <paramref name="match"/>.
        /// </summary>
        /// <param name="match">predicate to filter out matching validations</param>
        /// <returns></returns>
        public IEnumerable<ExcelDataValidation> FindAll(Predicate<ExcelDataValidation> match)
        {
            return _validations.FindAll(match);
        }

        /// <summary>
        /// Removes an <see cref="ExcelDataValidation"/> from the collection.
        /// </summary>
        /// <param name="item">The item to remove</param>
        /// <returns>True if remove succeeds, otherwise false</returns>
        /// <exception cref="ArgumentNullException">if <paramref name="item"/> is null</exception>
        public bool Remove(ExcelDataValidation item)
        {
            Require.Argument(item).IsNotNull("item");
            if (!(item is ExcelDataValidation))
            {
                throw new InvalidCastException("The supplied item must inherit OfficeOpenXml.DataValidation.ExcelDataValidation");
            }

            var retVal = _validations.Remove(item);
            if (retVal) OnValidationCountChanged();
            return retVal;
        }

        /// <summary>
        /// Returns the first matching validation.
        /// </summary>
        /// <param name="match"></param>
        /// <returns></returns>
        public ExcelDataValidation Find(Predicate<ExcelDataValidation> match)
        {
            return _validations.Find(match);
        }

        /// <summary>
        /// Removes all validations from the collection.
        /// </summary>
        public void Clear()
        {
            _validations.Clear();
        }

        /// <summary>
        /// Removes the validations that matches the predicate
        /// </summary>
        /// <param name="match"></param>
        public void RemoveAll(Predicate<ExcelDataValidation> match)
        {
            var matches = _validations.FindAll(match);
            foreach (var m in matches)
            {
                if (!(m is ExcelDataValidation))
                {
                    throw new InvalidCastException("The supplied item must inherit OfficeOpenXml.DataValidation.ExcelDataValidation");
                }
            }
            _validations.RemoveAll(match);
        }


        IEnumerator<ExcelDataValidation> IEnumerable<ExcelDataValidation>.GetEnumerator()
        {
            return _validations.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _validations.GetEnumerator();
        }

        private List<ExcelDataValidation> GetValidations()
        {
            return _validations;
        }

        internal void InsertRangeDictionary(ExcelAddress address, bool shiftRight)
        {
            if (address.Addresses == null)
            {
                InsertRangeIntoRangeDictionary(address, shiftRight);
            }
            else
            {
                foreach (var a in address.Addresses)
                {
                    InsertRangeIntoRangeDictionary(a, shiftRight);
                }
            }
        }
        private void InsertRangeIntoRangeDictionary(ExcelAddress address, bool shiftRight)
        {
            if (shiftRight)
            {
                _validationsRD.InsertColumn(address._fromCol, address.Columns, address._fromRow, address._toRow);
            }
            else
            {
                _validationsRD.InsertRow(address._fromRow, address.Rows, address._fromCol, address._toCol);
            }
        }

        internal void DeleteRangeDictionary(ExcelAddress address, bool shiftLeft)
        {
            if (address.Addresses == null)
            {
                DeleteRangeInRangeDictionary(address, shiftLeft);
            }
            else
            {
                foreach (var a in address.Addresses)
                {
                    DeleteRangeInRangeDictionary(a, shiftLeft);
                }
            }
        }
        private void DeleteRangeInRangeDictionary(ExcelAddress address, bool shiftLeft)
        {
            if (shiftLeft)
            {
                _validationsRD.DeleteColumn(address._fromCol, address.Columns, address._fromRow, address._toRow);
            }
            else
            {
                _validationsRD.DeleteRow(address._fromRow, address.Rows, address._fromCol, address._toCol);
            }
        }
    }
}
