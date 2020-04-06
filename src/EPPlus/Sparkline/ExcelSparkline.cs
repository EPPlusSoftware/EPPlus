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
using System.Xml;

namespace OfficeOpenXml.Sparkline
{
    /// <summary>
    /// Represents a single sparkline within the sparkline group
    /// </summary>
    public class ExcelSparkline : XmlHelper
    {
        internal ExcelSparkline(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
            SchemaNodeOrder = new string[] { "f", "sqref" };
        }   
        const string _fPath = "xm:f";
        /// <summary>
        /// The datarange
        /// </summary>
        public ExcelAddressBase RangeAddress
        {
            get
            {
                var v = GetXmlNodeString(_fPath);
                if(string.IsNullOrEmpty(v))
                {
                    return null;
                }
                else
                {
                    return new ExcelAddressBase(v);
                }
            }
            internal set
            {
                if(value==null || value.Address=="#REF!")
                {
                    DeleteNode(_fPath);
                }
                else
                {
                    SetXmlNodeString(_fPath, value.FullAddress);
                }
            }
        }
        const string _sqrefPath = "xm:sqref";
        /// <summary>
        /// Location of the sparkline
        /// </summary>
        public ExcelCellAddress Cell
        {
            get
            {
                return new ExcelCellAddress(GetXmlNodeString(_sqrefPath));
            }
            internal set
            {
                SetXmlNodeString("xm:sqref", value.Address);
            }
        }
        /// <summary>
        /// Returns a string representation of the object
        /// </summary>
        /// <returns>The cell address and the range</returns>
        public override string ToString()
        {
            return Cell.Address + ", " +RangeAddress.Address;
        }
    }
}
