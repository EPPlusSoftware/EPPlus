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
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    /// <summary>
    /// Adress over a range
    /// </summary>
    public class RangeAddress
    {
        /// <summary>
        /// Constructor for empty address
        /// </summary>
        public RangeAddress()
        {
            Address = string.Empty;
        }

        internal string Address { get; set; }
        /// <summary>
        /// Worksheet
        /// </summary>
        public string Worksheet { get; internal set; }
        /// <summary>
        /// From Column
        /// </summary>
        public int FromCol { get; internal set; }
        /// <summary>
        /// To Column
        /// </summary>
        public int ToCol { get; internal set; }
        /// <summary>
        /// From row
        /// </summary>
        public int FromRow { get; internal set; }
        /// <summary>
        /// To row
        /// </summary>
        public int ToRow { get; internal set; }
        /// <summary>
        /// To string
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Address;
        }

        private static RangeAddress _empty = new RangeAddress();
        /// <summary>
        /// Empty
        /// </summary>
        public static RangeAddress Empty
        {
            get { return _empty; }
        }

        /// <summary>
        /// Returns true if this range collides (full or partly) with the supplied range
        /// </summary>
        /// <param name="other">The range to check</param>
        /// <returns></returns>
        public bool CollidesWith(RangeAddress other)
        {
            if (other.Worksheet != Worksheet)
            {
                return false;
            }
            if (other.FromRow > ToRow || other.FromCol > ToCol
                ||
                FromRow > other.ToRow || FromCol > other.ToCol)
            {
                return false;
            }
            return true;
        }
    }
}
