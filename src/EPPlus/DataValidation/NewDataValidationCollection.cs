
using System;
using System.Collections;
using System.Collections.Generic;


using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Utils;
using OfficeOpenXml.DataValidation.Events;
using OfficeOpenXml.DataValidation.Formulas;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Math.RoundingHelper;

namespace OfficeOpenXml.DataValidation
{
    public class NewDataValidationCollection : IEnumerable<IExcelDataValidation>
    {
        private List<IExcelDataValidation> _validations = new List<IExcelDataValidation>();
        private ExcelWorksheet _worksheet = null;

        private const string DataValidationPath = "//d:dataValidations";
        private readonly string DataValidationItemsPath = string.Format("{0}/d:dataValidation", DataValidationPath);

        T AddValidation<T>(string address, ExcelDataValidationType ValidationType, Type type)
            where T : IExcelDataValidation
        {
            //ValidateAddress(address);
            //EnsureRootElementExists();
            // Object item = new ExcelDataValidationAny(_worksheet, ExcelDataValidation.NewId(), address, ValidationType);
            //_validations.Add(item);
            //OnValidationCountChanged();

            ValidateAddress(address);
            EnsureRootElementExists();
            Object item = Activator.CreateInstance(type, _worksheet, ExcelDataValidation.NewId(), ValidationType);
            _validations.Add((T)item);
            OnValidationCountChanged();
            //Object item = Activator.CreateInstance(typeof(T));    
            return (T)item;
        }

        private void OnValidationCountChanged()
        {

            ////if (TopNode != null)
            ////{
            ////    SetXmlNodeString("@count", _validations.Count.ToString());
            ////}
            //var dvNode = GetRootNode();
            //if (_validations.Count == 0)
            //{
            //    if (dvNode != null)
            //    {
            //        _worksheet.WorksheetXml.DocumentElement.RemoveChild(dvNode);
            //        TopNode = _worksheet.WorksheetXml.DocumentElement;
            //    }
            //    //_worksheet.ClearValidations();
            //}
            //else
            //{
            //    var attr = _worksheet.WorksheetXml.DocumentElement.SelectSingleNode(DataValidationPath + "[@count]", _worksheet.NameSpaceManager);
            //    if (attr == null)
            //    {
            //        dvNode.Attributes.Append(_worksheet.WorksheetXml.CreateAttribute("count"));
            //    }
            //    dvNode.Attributes["count"].Value = _validations.Count.ToString(CultureInfo.InvariantCulture);
            //}
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

        private void ValidateAddress(string address)
        {
            ValidateAddress(address, null);
        }

        private void EnsureRootElementExists()
        {
            //var node = _worksheet.WorksheetXml.SelectSingleNode(DataValidationPath, _worksheet.NameSpaceManager);
            //if (node == null)
            //{
            //    CreateNode(DataValidationPath.TrimStart('/'));
            //}
        }



        //T AddValidationOfType(Type type)
        //{

        //}


        //public IExcelDataValidationAny AddAnyValidation(string address, Action<ExcelDataValidationAny> action)
        //{
        //    var item = new ExcelDataValidationAny(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.Any);
        //    action.Invoke(item);
        //    return item;
        //}

        /// <summary>
        /// Adds a <see cref="ExcelDataValidationAny"/> to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public IExcelDataValidationAny AddAnyValidation(string address) => 
            AddValidation<IExcelDataValidationAny>(address, ExcelDataValidationType.Any, typeof(ExcelDataValidationAny));
     

        ///// <summary>
        ///// Adds an <see cref="IExcelDataValidationInt"/> to the worksheet. Whole means that the only accepted values
        ///// are integer values.
        ///// </summary>
        ///// <param name="address">the range/address to validate</param>
        //public IExcelDataValidationInt AddIntegerValidation(string address)
        //{
        //    ValidateAddress(address);
        //    EnsureRootElementExists();
        //    var item = new ExcelDataValidationInt(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.Whole);
        //    _validations.Add(item);
        //    OnValidationCountChanged();
        //    return item;
        //}

        ///// <summary>
        ///// Addes an <see cref="IExcelDataValidationDecimal"/> to the worksheet. The only accepted values are
        ///// decimal values.
        ///// </summary>
        ///// <param name="address">The range/address to validate</param>
        ///// <returns></returns>
        //public IExcelDataValidationDecimal AddDecimalValidation(string address)
        //{
        //    ValidateAddress(address);
        //    EnsureRootElementExists();
        //    var item = new ExcelDataValidationDecimal(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.Decimal);
        //    _validations.Add(item);
        //    OnValidationCountChanged();
        //    return item;
        //}

        ///// <summary>
        ///// Adds an <see cref="IExcelDataValidationList"/> to the worksheet. The accepted values are defined
        ///// in a list.
        ///// </summary>
        ///// <param name="address">The range/address to validate</param>
        ///// <returns></returns>
        //public IExcelDataValidationList AddListValidation(string address)
        //{
        //    ValidateAddress(address);
        //    EnsureRootElementExists();
        //    var item = new ExcelDataValidationList(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.List);
        //    ((ExcelDataValidationFormula)item.Formula).RegisterFormulaListener(_formulaListener);
        //    _validations.Add(item);
        //    OnValidationCountChanged();
        //    return item;
        //}

        //public IExcelDataValidationInt AddTextLengthValidation(string address)
        //{
        //    ValidateAddress(address);
        //    EnsureRootElementExists();
        //    var item = new ExcelDataValidationInt(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.TextLength);
        //    _validations.Add(item);
        //    OnValidationCountChanged();
        //    return item;
        //}

        ///// <summary>
        ///// Adds an <see cref="IExcelDataValidationDateTime"/> to the worksheet.
        ///// </summary>
        ///// <param name="address">The range/address to validate</param>
        ///// <returns></returns>
        //public IExcelDataValidationDateTime AddDateTimeValidation(string address)
        //{
        //    ValidateAddress(address);
        //    EnsureRootElementExists();
        //    var item = new ExcelDataValidationDateTime(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.DateTime);
        //    _validations.Add(item);
        //    OnValidationCountChanged();
        //    return item;
        //}


        ///// <summary>
        ///// Addes a <see cref="IExcelDataValidationTime"/> to the worksheet
        ///// </summary>
        ///// <param name="address">The range/address to validate</param>
        ///// <returns></returns>
        //public IExcelDataValidationTime AddTimeValidation(string address)
        //{
        //    ValidateAddress(address);
        //    EnsureRootElementExists();
        //    var item = new ExcelDataValidationTime(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.Time);
        //    _validations.Add(item);
        //    OnValidationCountChanged();
        //    return item;
        //}
        ///// <summary>
        ///// Adds a <see cref="ExcelDataValidationCustom"/> to the worksheet.
        ///// </summary>
        ///// <param name="address">The range/address to validate</param>
        ///// <returns></returns>
        //public IExcelDataValidationCustom AddCustomValidation(string address)
        //{
        //    ValidateAddress(address);
        //    EnsureRootElementExists();
        //    var item = new ExcelDataValidationCustom(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.Custom);
        //    ((ExcelDataValidationFormula)item.Formula).RegisterFormulaListener(_formulaListener);
        //    _validations.Add(item);
        //    OnValidationCountChanged();
        //    return item;
        //}

        //public bool Remove(IExcelDataValidation item)
        //{
        //    Require.Argument(item).IsNotNull("item");
        //    if (!(item is ExcelDataValidation))
        //    {
        //        throw new InvalidCastException("The supplied item must inherit OfficeOpenXml.DataValidation.ExcelDataValidation");
        //    }

        //    ((ExcelDataValidation)item).Delete();
        //    var retVal = _validations.Remove(item);
        //    if (retVal) OnValidationCountChanged();
        //    return retVal;
        //}

        /// <summary>
        /// Number of validations
        /// </summary>
        public int Count
        {
            get { return GetValidations().Count; }
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
        public IExcelDataValidation this[int index]
        {
            get { return GetValidations()[index]; }
            set { GetValidations()[index] = value; }
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
                return GetValidations().Find(x => x.Address.Collide(searchedAddress) != ExcelAddressBase.eAddressCollition.No);
            }
        }

        /// <summary>
        /// Returns all validations that matches the supplied predicate <paramref name="match"/>.
        /// </summary>
        /// <param name="match">predicate to filter out matching validations</param>
        /// <returns></returns>
        public IEnumerable<IExcelDataValidation> FindAll(Predicate<IExcelDataValidation> match)
        {
            return GetValidations().FindAll(match);
        }

        /// <summary>
        /// Returns the first matching validation.
        /// </summary>
        /// <param name="match"></param>
        /// <returns></returns>
        public IExcelDataValidation Find(Predicate<IExcelDataValidation> match)
        {
            return GetValidations().Find(match);
        }

        /// <summary>
        /// Removes all validations from the collection.
        /// </summary>
        public void Clear()
        {
            //if (TopNode != null && !string.IsNullOrEmpty(TopNode.LocalName) && TopNode.LocalName.ToLower() == "datavalidations")
            //{
            //    TopNode.ParentNode.RemoveChild(TopNode);
            //}
            //_validations.Clear();
            //_extLstValidations.Clear();
            //OnValidationCountChanged();
        }

        /// <summary>
        /// Removes the validations that matches the predicate
        /// </summary>
        /// <param name="match"></param>
        public void RemoveAll(Predicate<IExcelDataValidation> match)
        {
            //var matches = _validations.FindAll(match);
            //foreach (var m in matches)
            //{
            //    if (!(m is ExcelDataValidation))
            //    {
            //        throw new InvalidCastException("The supplied item must inherit OfficeOpenXml.DataValidation.ExcelDataValidation");
            //    }
            //    TopNode.RemoveChild(((ExcelDataValidation)m).TopNode);
            //    //var dvNode = TopNode.SelectSingleNode(DataValidationPath.TrimStart('/'), NameSpaceManager);
            //    //if (dvNode != null)
            //    //{
            //    //    dvNode.RemoveChild(((ExcelDataValidation)m).TopNode);
            //    //}
            //}
            //_validations.RemoveAll(match);
            //OnValidationCountChanged();
        }


        IEnumerator<IExcelDataValidation> IEnumerable<IExcelDataValidation>.GetEnumerator()
        {
            return GetValidations().GetEnumerator();
        }

        IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetValidations().GetEnumerator();
        }

        private List<IExcelDataValidation> GetValidations()
        {
            //if (_extLstValidations != null)
            //{
            //    var totalValidations = new List<IExcelDataValidation>(_validations);
            //    totalValidations.AddRange(_extLstValidations);
            //    return totalValidations;
            //}
            return _validations;
        }
    }
}
