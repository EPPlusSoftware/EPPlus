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
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml
{
    /// <summary>
    /// Range address with the address property readonly
    /// </summary>
    public class ExcelAddress : ExcelAddressBase
    {
        internal ExcelAddress()
            : base()
        {

        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fromRow">From row</param>
        /// <param name="fromCol">From column</param>
        /// <param name="toRow">To row</param>
        /// <param name="toColumn">To column</param>
        public ExcelAddress(int fromRow, int fromCol, int toRow, int toColumn)
            : base(fromRow, fromCol, toRow, toColumn)
        {
            _ws = "";
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="ws">Worksheet name</param>
        /// <param name="fromRow">From row</param>
        /// <param name="fromCol">From column</param>
        /// <param name="toRow">To row</param>
        /// <param name="toColumn">To column</param>
        public ExcelAddress(string ws, int fromRow, int fromCol, int toRow, int toColumn)
            : base(fromRow, fromCol, toRow, toColumn)
        {
            _ws = ws;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="address">The Excel address</param>
        public ExcelAddress(string address)
            : base(address)
        {
        }
        
        internal ExcelAddress(string ws, string address)
            : base(address)
        {
            if (string.IsNullOrEmpty(_ws)) _ws = ws;
        }
        internal ExcelAddress(string ws, string address, bool isName)
            : base(address, isName)
        {
            if (string.IsNullOrEmpty(_ws)) _ws = ws;
        }

        /// <summary>
        /// Creates an Address object
        /// </summary>
        /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
        /// <param name="Address">The Excel Address</param>
        /// <param name="package">Reference to the package to find information about tables and names</param>
        /// <param name="referenceAddress">The address</param>
        public ExcelAddress(string Address, ExcelPackage package, ExcelAddressBase referenceAddress) :
            base(Address, package, referenceAddress)
        {

        }
        /// <summary>
        /// The address for the range
        /// </summary>
        /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
        public new string Address
        {
            get
            {
                if (string.IsNullOrEmpty(_address) && _fromRow>0)
                {
                    _address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
                }
                return _address;
            }
            set
            {                
                SetAddress(value, null, null);
                ChangeAddress();
            }
        }
    }
}
