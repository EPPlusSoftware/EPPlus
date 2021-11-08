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
using System.Collections;
using System.Globalization;
using OfficeOpenXml.Utils;
using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas;

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
    public class ExcelDataValidationCollection : XmlHelper, IEnumerable<IExcelDataValidation>
    {
        private List<IExcelDataValidation> _validations = new List<IExcelDataValidation>();
        private ExcelExLstDataValidationCollection _extLstValidations = null;
        private ExcelWorksheet _worksheet = null;
        private readonly bool _extListUsed = false;
        private readonly DataValidationFormulaListener _formulaListener = null;

        private const string DataValidationPath = "//d:dataValidations";
        private readonly string DataValidationItemsPath = string.Format("{0}/d:dataValidation", DataValidationPath);

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        internal ExcelDataValidationCollection(ExcelWorksheet worksheet)
            : base(worksheet.NameSpaceManager, worksheet.WorksheetXml.DocumentElement)
        {
            Require.Argument(worksheet).IsNotNull("worksheet");
            _worksheet = worksheet;
            _formulaListener = new DataValidationFormulaListener(this, _worksheet);
            SchemaNodeOrder = worksheet.SchemaNodeOrder;

            // check existing nodes and load them
            var dataValidationNodes = worksheet.WorksheetXml.SelectNodes(DataValidationItemsPath, worksheet.NameSpaceManager);
            if (dataValidationNodes != null && dataValidationNodes.Count > 0)
            {
                foreach (XmlNode node in dataValidationNodes)
                {
                    if (node.Attributes["sqref"] == null) continue;

                    var addr = node.Attributes["sqref"].Value;
                    var uid = node.Attributes["xr:uid"] != null && !string.IsNullOrEmpty(node.Attributes["xr:uid"].Value) ? node.Attributes["xr:uid"].Value : ExcelDataValidation.NewId();
                    var typeSchema = node.Attributes["type"] != null ? node.Attributes["type"].Value : "";

                    var type = ExcelDataValidationType.GetBySchemaName(typeSchema);
                    var validation = ExcelDataValidationFactory.Create(type, worksheet, addr, node, InternalValidationType.DataValidation, uid);
                    validation.RegisterFormulaListener(_formulaListener);
                    validation.Uid = uid;
                    _validations.Add(validation);
                }
            }
            if (_validations.Count > 0)
            {
                OnValidationCountChanged();
            }
            _extLstValidations = new ExcelExLstDataValidationCollection(worksheet, _formulaListener);
            _extListUsed = !_extLstValidations.IsEmpty;
            InternalValidationEnabled = true;

            if(worksheet.WorksheetXml.DocumentElement!=null)
            {
                var xr=worksheet.WorksheetXml.DocumentElement.GetAttribute("xmlns:xr");
                if(string.IsNullOrEmpty(xr))
                {
                    worksheet.WorksheetXml.DocumentElement.SetAttribute("xmlns:xr", ExcelPackage.schemaXr);
                    var mc = worksheet.WorksheetXml.DocumentElement.GetAttribute("xmlns:mc");
                    if(mc != ExcelPackage.schemaMarkupCompatibility)
                    {
                        worksheet.WorksheetXml.DocumentElement.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);
                    }
                    var ignore = worksheet.WorksheetXml.DocumentElement.GetAttribute("mc:Ignorable");
                    var nsIgnore = ignore.Split(' ');
                    if (!nsIgnore.Contains("xr"))
                    {
                        worksheet.WorksheetXml.DocumentElement.SetAttribute("Ignorable",ExcelPackage.schemaMarkupCompatibility, string.IsNullOrEmpty(ignore) ? "xr" : ignore + " xr");
                    }
                }
            }
        }

        internal void AddCopyOfDataValidation(string address, ExcelDataValidation dv)
        {
            EnsureRootElementExists();
            var node = CreateNode(DataValidationItemsPath, false,true);
            CopyElement((XmlElement)dv.TopNode, (XmlElement)node);
            var validation = ExcelDataValidationFactory.Create(dv.ValidationType, _worksheet, address, node, InternalValidationType.DataValidation, ExcelDataValidation.NewId());
            _validations.Add(validation);
        }

        private void EnsureRootElementExists()
        {
            var node = _worksheet.WorksheetXml.SelectSingleNode(DataValidationPath, _worksheet.NameSpaceManager);
            if (node == null)
            {
                CreateNode(DataValidationPath.TrimStart('/'));
            }
        }

        private void OnValidationCountChanged()
        {

            //if (TopNode != null)
            //{
            //    SetXmlNodeString("@count", _validations.Count.ToString());
            //}
            var dvNode = GetRootNode();
            if (_validations.Count == 0)
            {
                if (dvNode != null)
                {
                    _worksheet.WorksheetXml.DocumentElement.RemoveChild(dvNode);
                    TopNode = _worksheet.WorksheetXml.DocumentElement;
                }
                //_worksheet.ClearValidations();
            }
            else
            {
                var attr = _worksheet.WorksheetXml.DocumentElement.SelectSingleNode(DataValidationPath + "[@count]", _worksheet.NameSpaceManager);
                if (attr == null)
                {
                    dvNode.Attributes.Append(_worksheet.WorksheetXml.CreateAttribute("count"));
                }
                dvNode.Attributes["count"].Value = _validations.Count.ToString(CultureInfo.InvariantCulture);
            }
        }

        private XmlNode GetRootNode()
        {
            EnsureRootElementExists();
            TopNode = _worksheet.WorksheetXml.SelectSingleNode(DataValidationPath, _worksheet.NameSpaceManager);
            return TopNode;
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

        internal ExcelExLstDataValidationCollection DataValidationsExt
        {
            get { return _extLstValidations; }
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
            if(_extLstValidations != null)
            {
                foreach(var extValidation in _extLstValidations)
                {
                    extValidation.Validate();
                }
            }
        }

        /// <summary>
        /// Adds a <see cref="ExcelDataValidationAny"/> to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public IExcelDataValidationAny AddAnyValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationAny(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.Any);
            _validations.Add(item);
            OnValidationCountChanged();
            return item;
        }

        /// <summary>
        /// Adds an <see cref="IExcelDataValidationInt"/> to the worksheet. Whole means that the only accepted values
        /// are integer values.
        /// </summary>
        /// <param name="address">the range/address to validate</param>
        public IExcelDataValidationInt AddIntegerValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationInt(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.Whole);
            _validations.Add(item);
            OnValidationCountChanged();
            return item;
        }

        /// <summary>
        /// Addes an <see cref="IExcelDataValidationDecimal"/> to the worksheet. The only accepted values are
        /// decimal values.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public IExcelDataValidationDecimal AddDecimalValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationDecimal(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.Decimal);
            _validations.Add(item);
            OnValidationCountChanged();
            return item;
        }

        /// <summary>
        /// Adds an <see cref="IExcelDataValidationList"/> to the worksheet. The accepted values are defined
        /// in a list.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public IExcelDataValidationList AddListValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationList(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.List);
            ((ExcelDataValidationFormula)item.Formula).RegisterFormulaListener(_formulaListener);
            _validations.Add(item);
            OnValidationCountChanged();
            return item;
        }

        internal IExcelDataValidationList AddListValidation(string address, string uid)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationList(_worksheet, uid, address, ExcelDataValidationType.List);
            ((ExcelDataValidationFormula)item.Formula).RegisterFormulaListener(_formulaListener);
            _validations.Add(item);
            OnValidationCountChanged();
            return item;
        }

        /// <summary>
        /// Adds an <see cref="IExcelDataValidationInt"/> regarding text length to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public IExcelDataValidationInt AddTextLengthValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationInt(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.TextLength);
            _validations.Add(item);
            OnValidationCountChanged();
            return item;
        }

        /// <summary>
        /// Adds an <see cref="IExcelDataValidationDateTime"/> to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public IExcelDataValidationDateTime AddDateTimeValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationDateTime(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.DateTime);
            _validations.Add(item);
            OnValidationCountChanged();
            return item;
        }


        /// <summary>
        /// Addes a <see cref="IExcelDataValidationTime"/> to the worksheet
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public IExcelDataValidationTime AddTimeValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationTime(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.Time);
            _validations.Add(item);
            OnValidationCountChanged();
            return item;
        }
        /// <summary>
        /// Adds a <see cref="ExcelDataValidationCustom"/> to the worksheet.
        /// </summary>
        /// <param name="address">The range/address to validate</param>
        /// <returns></returns>
        public IExcelDataValidationCustom AddCustomValidation(string address)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationCustom(_worksheet, ExcelDataValidation.NewId(), address, ExcelDataValidationType.Custom);
            ((ExcelDataValidationFormula)item.Formula).RegisterFormulaListener(_formulaListener);
            _validations.Add(item);
            OnValidationCountChanged();
            return item;
        }

        internal IExcelDataValidationCustom AddCustomValidation(string address, string uid)
        {
            ValidateAddress(address);
            EnsureRootElementExists();
            var item = new ExcelDataValidationCustom(_worksheet, uid, address, ExcelDataValidationType.Custom);
            ((ExcelDataValidationFormula)item.Formula).RegisterFormulaListener(_formulaListener);
            _validations.Add(item);
            OnValidationCountChanged();
            return item;
        }

        /// <summary>
        /// Removes an <see cref="ExcelDataValidation"/> from the collection.
        /// </summary>
        /// <param name="item">The item to remove</param>
        /// <returns>True if remove succeeds, otherwise false</returns>
        /// <exception cref="ArgumentNullException">if <paramref name="item"/> is null</exception>
        public bool Remove(IExcelDataValidation item)
        {
            Require.Argument(item).IsNotNull("item");
            if (!(item is ExcelDataValidation))
            {
                throw new InvalidCastException("The supplied item must inherit OfficeOpenXml.DataValidation.ExcelDataValidation");
            }

            ((ExcelDataValidation)item).Delete();
            var retVal = _validations.Remove(item);
            if (retVal) OnValidationCountChanged();
            return retVal;
        }

        private List<IExcelDataValidation> GetValidations()
        {
            if(_extLstValidations != null)
            {
                var totalValidations = new List<IExcelDataValidation>(_validations);
                totalValidations.AddRange(_extLstValidations);
                return totalValidations;
            }
            return _validations;
        }

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
            DeleteAllNode(DataValidationItemsPath.TrimStart('/'));
            _validations.Clear();
            _extLstValidations.Clear();
        }

        /// <summary>
        /// Removes the validations that matches the predicate
        /// </summary>
        /// <param name="match"></param>
        public void RemoveAll(Predicate<IExcelDataValidation> match)
        {
            var matches = _validations.FindAll(match);
            foreach (var m in matches)
            {
                if (!(m is ExcelDataValidation))
                {
                    throw new InvalidCastException("The supplied item must inherit OfficeOpenXml.DataValidation.ExcelDataValidation");
                }
                TopNode.RemoveChild(((ExcelDataValidation)m).TopNode);
                //var dvNode = TopNode.SelectSingleNode(DataValidationPath.TrimStart('/'), NameSpaceManager);
                //if (dvNode != null)
                //{
                //    dvNode.RemoveChild(((ExcelDataValidation)m).TopNode);
                //}
            }
            _validations.RemoveAll(match);
            OnValidationCountChanged();
        }

        IEnumerator<IExcelDataValidation> IEnumerable<IExcelDataValidation>.GetEnumerator()
        {
            return GetValidations().GetEnumerator();
        }

        IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetValidations().GetEnumerator();
        }
    }
}
