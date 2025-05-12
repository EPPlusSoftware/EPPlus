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
using OfficeOpenXml.Core;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Utils.Formula;

namespace OfficeOpenXml
{
    /// <summary>
    /// A named range. 
    /// </summary>
    public sealed class ExcelNamedRange : ExcelRangeBase 
    {
        [Flags]
        private enum NameRelativeType
        {
            /// <summary>
            /// The name contains a relative address
            /// </summary>
            RelativeAddress = 1,
            /// <summary>
            /// The name contains a relative table address, i.e. [#this row]
            /// </summary>
            RelativeTableAddress = 2
        }
        ExcelWorksheet _sheet;
        /// <summary>
        /// A named range
        /// </summary>
        /// <param name="name">The name</param>
        /// <param name="nameSheet">The sheet containing the name. null if its a global name</param>
        /// <param name="sheet">Sheet where the address points</param>
        /// <param name="address">The address</param>
        /// <param name="index">The index in the collection</param>
        /// <param name="allowRelativeAddress">If true, the address will be retained as it is, if false the address will always be converted to an absolute/fixed address</param>
        internal ExcelNamedRange(string name, ExcelWorksheet nameSheet , ExcelWorksheet sheet, string address, int index, bool allowRelativeAddress = false) :
            base(sheet, address)
        {
            Init(name, nameSheet, index, allowRelativeAddress);
        }
        internal ExcelNamedRange(string name,ExcelWorkbook wb, ExcelWorksheet nameSheet, int index, bool allowRelativeAddress = false) :
            base(wb, nameSheet, name, true)
        {
            Init(name, nameSheet, index, allowRelativeAddress);
        }
        private void Init(string name, ExcelWorksheet nameSheet, int index, bool allowRelativeAddress)
        {
            Name = name;
            _sheet = nameSheet;
            Index = index;
            if(allowRelativeAddress && _fromRow>0)
            {                
                if(IsFullColumn)
                {
                    _relativeType = (_fromColFixed && _toColFixed) ? 0 : NameRelativeType.RelativeAddress;
                }
                else if(IsFullRow)
                {
                    _relativeType = (_fromRowFixed && _toRowFixed) ? 0 : NameRelativeType.RelativeAddress;
                }
                else
                {
                    _relativeType = (_fromRowFixed && _toRowFixed && _fromColFixed && _toColFixed) ? 0 : NameRelativeType.RelativeAddress;
                }
            }
            else if(_fromRow>0 && !(_fromRowFixed && _toRowFixed && _fromColFixed && _toColFixed))
            {
                _fromRowFixed = _toRowFixed = _fromColFixed = _toColFixed = true;
                ResetAddress(_address);
            }
        }

        /// <summary>
        /// Name of the range
        /// </summary>
        public string Name
        {
            get;
            internal set;
        }
        /// <summary>
        /// Is the named range local for the sheet 
        /// </summary>
        public int LocalSheetId
        {
            get
            {
                if (_sheet == null)
                {
                    return -1;
                }
                else
                {
                    return _sheet.IndexInList;
                }
            }
        }
        internal ExcelWorksheet LocalSheet => _sheet;

        internal int Index
        {
            get;
            set;
        }
        /// <summary>
        /// Is the name hidden
        /// </summary>
        public bool IsNameHidden
        {
            get;
            set;
        }
        /// <summary>
        /// A comment for the Name
        /// </summary>
        public string NameComment
        {
            get;
            set;
        }
        internal object NameValue 
        { 
            get; 
            set; 
        }
        IList<Token> _tokens = null;
        /// <summary>
        /// Set to true to validate and update formulas with cell references.
        /// </summary>
        public static bool ValidateCellAddressInFormulas = true;
        private string _nameFormula;
        internal string NameFormula
        {
            get
            {
                return _nameFormula;
            }
            set
            {
                if (value == null)
                {
                    _nameFormula = value;
                    return;
                }
                _nameFormula = ValidateCellAddressInFormulas ? FormulaUtils.AddWorksheetReferenceToFormula(value, _worksheet, AllowRelativeAddress) : value;
            }
        }
        /// <inheritdoc/>
        protected internal override void BeforeChangeAddress(string value)
        {
            if(!string.IsNullOrEmpty(value) && !value.Contains("!") && Worksheet == null)
            {
                throw new InvalidOperationException("Workbook name needs a worksheet in the address.");
            }

        }
        /// <inheritdoc/>
        protected internal override void ChangeAddress(string value)
        {

        }

        /// <summary>
        /// Set a range for this name. Will remove exsisting formula and value.
        /// </summary>
        /// <param name="range"></param>
        /// <param name="allowRelativeAddress"></param>
        public void SetRange(ExcelRangeBase range, bool allowRelativeAddress = false)
        {
            if (range.Worksheet != _worksheet)
            {
                throw new InvalidOperationException($"Cannot change range to another worksheet or set a range to a workbook name: {Name}. Either create a new name or move current name first.");
            }
            ResetObject();
            NameFormula = null;
            NameValue = null;

            _workbook = range._workbook;
            _worksheet = range._worksheet;
            if (allowRelativeAddress)
            {
                Address = range.FullAddress;
            }
            else
            {
                Address = range.FullAddressAbsolute;
            }
            Value = range.Value;
        }

        /// <summary>
        /// Set a formula for this name. Will remove exsisting range and value.
        /// </summary>
        /// <param name="formula">The formula for this name.</param>
        public void SetFormula(string formula)
        {
            ResetObject();
            NameFormula = formula;
            this.Formula = NameFormula;
            NameValue = null;
        }

        /// <summary>
        /// Set a value for this name. Will remove exsisting range and formula.
        /// </summary>
        /// <param name="value"></param>
        public void SetValue(object value)
        {
            ResetObject();
            NameFormula = null;
            NameValue = value;
            Value = value;
        }
        /// <summary>
        /// Move this defined name to a target worksheet.
        /// </summary>
        /// <param name="worksheet">Worksheet to move this name to.</param>
        /// /// <param name="name">Optional new name for the defined name.</param>
        /// <returns>This name.</returns>
        /// <exception cref="InvalidOperationException">If this name does not contain a formula, value or range, this exception occurs.</exception>
        public ExcelNamedRange Move(ExcelWorksheet worksheet, string name = null)
        {
            ExcelNamedRange enr = null;
            name = name == null ? Name : name;
            //Detect if formula, value or range
            if (NameFormula != null)
                enr = worksheet.Names.AddFormula(Name, NameFormula);
            else if (NameValue != null)
                enr = worksheet.Names.AddValue(Name, NameValue);
            else if (LocalAddress != "#REF!")
                enr = worksheet.Names.AddRange(Name, worksheet.Cells[Address]);
            if (enr == null) throw new InvalidOperationException($"No value, formula or address has been set for this name: {Name}");
            if(_worksheet != null)
                _worksheet.Names.Remove(Name);
            else if(_workbook != null)
                _workbook.Names.Remove(Name);
            else
                throw new InvalidOperationException($"No workbook or worksheet has been set for this name: {Name}");
            return enr;
        }
        /// <summary>
        /// Move this defined name to target workbook.
        /// </summary>
        /// <param name="workbook">Workbook to move this name to.</param>
        /// <param name="name">Optional new name for the defined name.</param>
        /// <returns>This name.</returns>
        /// <exception cref="InvalidOperationException">If this name does not contain a formula, value or range, this exception occurs.</exception>
        public ExcelNamedRange Move(ExcelWorkbook workbook, string name = null)
        {
            ExcelNamedRange enr = null;
            name = name == null ? Name : name;
            //Detect if formula, value or range
            if (NameFormula != null)
                enr = workbook.Names.AddFormula(Name, NameFormula);
            else if (NameValue != null)
                enr = workbook.Names.AddValue(Name, NameValue);
            else if (LocalAddress != "#REF!")
                enr = workbook.Names.AddRange(Name, _worksheet.Cells[Address]);
            if(enr == null) throw new InvalidOperationException($"No value, formula or address has been set for this name: {Name}");
            if (_worksheet != null)
                _worksheet.Names.Remove(Name);
            else if (_workbook != null)
                _workbook.Names.Remove(Name);
            else
                throw new InvalidOperationException($"No workbook or worksheet has been set for this name: {Name}");
            return enr;
        }
        /// <summary>
        /// Creates a copy of the defined name to target worksheet.
        /// </summary>
        /// <param name="worksheet">Worksheet to copy to.</param>
        /// <param name="name">The name for the copy.</param>
        /// <returns>A new ExcelNameRange</returns>
        /// <exception cref="InvalidOperationException">If the original does not contain a formula, value or range, this exception occurs.</exception>
        public ExcelNamedRange Copy(ExcelWorksheet worksheet, string name)
        {
            //Detect if formula, value or range
            if (NameFormula != null)
                return worksheet.Names.AddFormula(name, NameFormula);
            else if (NameValue != null)
                return worksheet.Names.AddValue(name, NameValue);
            else if (LocalAddress != "#REF!")
                return worksheet.Names.AddRange(name, worksheet.Cells[Address]);
            throw new InvalidOperationException($"No value, formula or address has been set for this name: {Name}");
        }
        /// <summary>
        /// Creates a copy of the defined name to target workbook.
        /// </summary>
        /// <param name="workbook">Workbook to copy to.</param>
        /// <param name="name">The name for the copy.</param>
        /// <returns>A new ExcelNameRange</returns>
        /// <exception cref="InvalidOperationException">If the original does not contain a formula, value or range, this exception occurs.</exception>
        public ExcelNamedRange Copy(ExcelWorkbook workbook, string name)
        {
            //Detect if formula, value or range
            if (NameFormula != null)
                return workbook.Names.AddFormula(name, NameFormula);
            else if (NameValue != null)
                return workbook.Names.AddValue(name, NameValue);
            else if (LocalAddress != "#REF!")
                return workbook.Names.AddRange(name, _worksheet.Cells[Address]);
            throw new InvalidOperationException($"No value, formula or address has been set for this name: {Name}");
        }

        private void ResetObject()
        {
            Init(Name, _sheet, Index, AllowRelativeAddress);
            _address = Name;
            _fromCol = -1;
            _fromRow = -1;
            _toCol = -1;
            _toRow = -1;
            _start = null;
            _end = null;
        }

        string _r1c1Formula = "";
        internal string GetRelativeFormula(int row, int col)
        {
            if (string.IsNullOrEmpty(_r1c1Formula) && !string.IsNullOrEmpty(NameFormula))
            {
                if (_relativeType == NameRelativeType.RelativeTableAddress) return NameFormula;

                if (_tokens == null)
                {
                    SetRelativeType();
                }
                if (_relativeType == 0) return NameFormula;
                if((_relativeType & (NameRelativeType.RelativeAddress | NameRelativeType.RelativeTableAddress)) != 0)                   
                {
                    _r1c1Formula = R1C1Translator.ToR1C1FromTokens(_tokens, 1, 1);
                }
                else
                {
                    _r1c1Formula = NameFormula;
                }
            }
            else if(IsRelative == false) 
            {
                return NameFormula;
            }

            return GetRelativeFormula(_r1c1Formula, row, col);
        }

        private void SetRelativeType()
        {
            _tokens = SourceCodeTokenizer.Default.Tokenize(NameFormula);
            _relativeType = HasRelativeRef(_tokens);
        }

        private NameRelativeType HasRelativeRef(IList<Token> tokens)
        {
            NameRelativeType ret=0;
            foreach(var t in tokens)
            {
                if(t.TokenType==TokenType.CellAddress)
                {
                    if(t.Value.Count(x=>x=='$')<2)
                    {
                        ret |= NameRelativeType.RelativeAddress;
                    }
                }
                else if(t.TokenType==TokenType.TablePart)
                {
                    if (t.Value.Equals("#this row", StringComparison.InvariantCultureIgnoreCase))
                    {
                        ret |= NameRelativeType.RelativeTableAddress;
                    }
                }
            }
            return ret;
        }

        internal static string GetRelativeFormula(string sourceFormula, int row, int col)
        {
            var formula = "";
            var tokens = SourceCodeTokenizer.R1C1.Tokenize(sourceFormula);
            foreach (var t in tokens)
            {
                switch (t.TokenType)
                {
                    case TokenType.ExcelAddressR1C1:
                        formula += R1C1Translator.FromR1C1Formula(t.Value, row, col, true);
                        break;
                    default:
                        formula += t.Value;
                        break;
                }
            }
            return formula;
        }

        private void GetRefAddress(out int row, out int col)
        {
            var ix = _workbook.View.ActiveTab;
            var activeCell =  _workbook.GetWorksheetByIndexInList(ix).View.ActiveCell;
            ExcelCellBase.GetRowCol(activeCell, out row, out col, false);
        }
        /// <summary>
        /// Returns a string representation of the object
        /// </summary>
        /// <returns>The name of the range</returns>
        public override string ToString()
        {
            return Name;
        }
        /// <summary>
        /// Returns true if the name is equal to the obj
        /// </summary>
        /// <param name="obj">The object to compare with</param>
        /// <returns>true if equal</returns>
        public override bool Equals(object obj)
        {
            if(obj is ExcelNamedRange name)
            {
                return name.Name.Equals(Name, StringComparison.OrdinalIgnoreCase) && 
                       name.LocalSheetId == LocalSheetId && 
                       name._workbook == _workbook;
            }
            else
            {
                return base.Equals(obj);
            }
        }

        /// <summary>
        ///  If true, the address will be retained as it is, if false the address will always be converted to an absolute/fixed address
        /// </summary>
        internal bool AllowRelativeAddress
        {
            get; private set;
        }

        NameRelativeType _relativeType;
        internal bool IsRelative
        {
            get
            {
                if(!string.IsNullOrEmpty(NameFormula) && _tokens==null)
                {
                    SetRelativeType();
                    AllowRelativeAddress = _relativeType > 0;
                }
                return _relativeType > 0;
            }
        }
        /// <summary>
        /// Serves as the default hash function.
        /// </summary>
        /// <returns>A hash code for the current object.</returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        internal object GetValue(FormulaCellAddress currentCell)
        {
            if (IsRelative)
            {
                if(string.IsNullOrEmpty(NameFormula))
                {
                    var ri = NameValue as RangeInfo;
                    if (ri == null) 
                    {
                        return NameValue;
                    }
                    else
                    {
                        return GetRelativeRange(ri, currentCell);
                    }
                }
                else
                {
                    var values = NameValue as Dictionary<ulong, object>;
                    if(values!=null)
                    {
                        if(values.ContainsKey(currentCell.CellId))
                        {
                            return values[currentCell.CellId];
                        }
                        else
                        {
                            return null;
                        }
                        
                    }
                    return NameValue;
                }
            }
            else
            {
                return NameValue;
            }
        }

        internal RangeInfo GetRelativeRange(IRangeInfo ri, FormulaCellAddress currentCell)
        {
            var address = ri.Address.GetOffset(currentCell.Row, currentCell.Column, true);
            return new RangeInfo(address);
        }

        internal void SetValue(object resultValue, FormulaCellAddress currentCell)
        {
            if(AllowRelativeAddress)
            {
                Dictionary<ulong, object> values;
                if(NameValue==null)
                {
                    values = new Dictionary<ulong, object>();
                    NameValue = values; 
                }
                else
                {
                    values = (Dictionary<ulong, object>)NameValue;
                }
                if (values.ContainsKey(currentCell.CellId) == false)
                {
                    values.Add(currentCell.CellId, resultValue);
                }
            }
            else
            {
                NameValue = resultValue;
            }
        }
    }
}
