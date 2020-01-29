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

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A field Item. Used for grouping
    /// </summary>
    public class ExcelPivotTableFieldItem : XmlHelper
    {
        ExcelPivotTableField _field;
        internal ExcelPivotTableFieldItem(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field) :
            base(ns, topNode)
        {
           _field = field;
        }
        /// <summary>
        /// The text. Unique values only
        /// </summary>
        public string Text
        {
            get
            {
                return GetXmlNodeString("@n");
            }
            set
            {
                if(string.IsNullOrEmpty(value))
                {
                    DeleteNode("@n");
                    return;
                }
                foreach (var item in _field.Items)
                {
                    if (item.Text == value)
                    {
                        throw(new ArgumentException("Duplicate Text"));
                    }
                }
                SetXmlNodeString("@n", value);
            }
        }
        internal int X
        {
            get
            {
                return GetXmlNodeInt("@x"); 
            }
        }
        internal string T
        {
            get
            {
                return GetXmlNodeString("@t");
            }
        }
    }
}
