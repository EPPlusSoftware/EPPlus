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
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// A named range. 
    /// </summary>
    public sealed class ExcelNamedRange : ExcelRangeBase 
    {
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
            Name = name;
            _sheet = nameSheet;
            Index = index;
            AllowRelativeAddress = allowRelativeAddress;
        }
        internal ExcelNamedRange(string name,ExcelWorkbook wb, ExcelWorksheet nameSheet, int index, bool allowRelativeAddress = false) :
            base(wb, nameSheet, name, true)
        {
            Name = name;
            _sheet = nameSheet;
            Index = index;
            AllowRelativeAddress = allowRelativeAddress;
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
        internal object NameValue { get; set; }
        internal string NameFormula { get; set; }
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

        /// <summary>
        /// Serves as the default hash function.
        /// </summary>
        /// <returns>A hash code for the current object.</returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
