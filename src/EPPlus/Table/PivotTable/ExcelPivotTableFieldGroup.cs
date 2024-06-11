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
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Base class for pivot table field groups
    /// </summary>
    public class ExcelPivotTableFieldGroup : XmlHelper
    {
        internal ExcelPivotTableFieldGroup(XmlNamespaceManager ns, XmlNode topNode) :
            base(ns, topNode)
        {
            
        }
        /// <summary>
        /// The index of the base field for this group field
        /// </summary>
        public int? BaseIndex
        {
            get
            {
                return GetXmlNodeIntNull("@base");
            }
            internal set
            {
                SetXmlNodeInt("@base", value);
            }
        }
        /// <summary>
        /// The index of the parent field from this group field
        /// </summary>
        public int? ParentIndex
        {
            get
            {
                return GetXmlNodeIntNull("@par");
            }
        }
    }
}
