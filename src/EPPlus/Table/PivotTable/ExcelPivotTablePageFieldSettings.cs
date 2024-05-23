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
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A page / report filter field
    /// </summary>
    public class ExcelPivotTablePageFieldSettings  : XmlHelper
    {
        internal ExcelPivotTableField _field;
        internal ExcelPivotTablePageFieldSettings(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field, int index) :
            base(ns, topNode)
        {
            if (GetXmlNodeString("@hier")=="")
            {
                Hier = -1;
            }
            _field = field;
        }
        internal int Index 
        { 
            get
            {
                return GetXmlNodeInt("@fld");
            }
            set
            {
                SetXmlNodeString("@fld",value.ToString());
            }
        }
        /// <summary>
        /// The display name of the hierarchy.
        /// </summary>
        public int Caption
        {
            get
            {
                return GetXmlNodeInt("@cap");
            }
            set
            {
                SetXmlNodeString("@cap", value.ToString());
            }
        }

        /// <summary>
        /// The Name of the field
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
            set
            {
                SetXmlNodeString("@name", value);
            }
        }
        /// <summary>
        /// The selected item. 
        /// See also <seealso cref="ExcelPivotTableFieldItemsCollection.SelectSingleItem(int)"/>
        /// </summary>
        public int SelectedItem
        {
            get
            {
                return GetXmlNodeInt("@item");
            }
            internal set
            {
                if (value < 0)
                {
                    DeleteNode("@item");
                }
                else
                {
                    SetXmlNodeString("@item", value.ToString());
                }
            }
        }
        /// <summary>
        /// The index of the OLAP hierarchy to which this page field belongs
        /// </summary>
        internal int Hier
        {
            get
            {
                return GetXmlNodeInt("@hier");
            }
            set
            {
                SetXmlNodeString("@hier", value.ToString());
            }
        }
    }
}
