/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  7/11/2023         EPPlus Software AB       EPPlus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Attributes
{
    /// <summary>
    /// This attributes can only be used on properties that are of the type IDictionary&lt;string, string&gt;.
    /// Columns will be added based on the items in <see cref="EPPlusDictionaryColumnAttribute.ColumnHeaders"/>
    /// </summary>
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Property | AttributeTargets.Field)]
    internal class EPPlusDictionaryColumnAttribute : Attribute
    {
        /// <summary>
        /// Order of the columns value, default value is 0
        /// </summary>
        public int Order
        {
            get;
            set;
        }

        /// <summary>
        /// The values of this array will be used to generate columns (one column for each item).
        /// </summary>
        public string[] ColumnHeaders { get; set; }

        /// <summary>
        /// Should be unique within all attributes. Will be used to retrieve the keys of the Dictionary
        /// that also will be used to create the columns for this property.
        /// </summary>
        public string KeyId { get; set; }
    }
}
